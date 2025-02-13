# annict2fani.ps1
# Annict.comから自分の書いたレビュー、評価、各回コメント、作品メモを全て取得し、デスクトップにFani通調査票形式CSV/Excelを出力
# Fani通のレビュー用に評点、視聴状況を適宜変換
#
# 注意
#  -dumpを付けて実行すると取得したJSONをファイル保存
#  各回コメントはレビューの本文に話数:コメントとしてマージ
#  複数回Revewがある場合は別行として出力(評価が違うと本文のみマージだと困るので)
#  レビューが無く各回コメントか作品メモのみがある場合は記入可能項目のみ追加(評点等は無し)

#初回実行前に個人用アクセストークンを以下から新規作成
# https://annict.com/settings/apps
#ユーザー環境変数 ANNICT_ACCESS_TOKEN に設定
$accessToken = $Env:ANNICT_ACCESS_TOKEN

#出力ファイル
$outputFilePath = Join-Path -Path ([System.Environment]::GetFolderPath("Desktop")) `
                    -ChildPath "annict_personal_review_$(get-date -Format "yyyyMMdd_HHmm")"

# Annict GraphQLエンドポイントURI
$endpoint = "https://api.annict.com/graphql"

#ユーザー名取得クエリ
$queryName = @"
query {
    viewer {
        name
    }
}
"@

#Reviewデータ取得用GraphQLクエリ(番組レビュー、評価取得)
$queryReviews = @"
query {
    viewer {
        activities {
            edges {
                item {
                    ... on Review {
                        work {
                            annictId
                            # syobocalTid
                            title
                            titleKana
                            seasonYear
                            seasonName
                            # started_on
                            episodesCount
                            viewerStatusState
                        }
                        body
                        ratingAnimationState
                        ratingCharacterState
                        ratingMusicState
                        ratingOverallState
                        ratingStoryState
                        updatedAt
                    }
                }
            }
        }
    }
}
"@

#Recordデータ取得用GraphQLクエリ(各回コメント取得)
$queryRecords = @'
query ($user: String!)  {
    user (username: $user)  {
        records	(hasComment: true) {
            nodes {
                work {
                    annictId
                    title
                    titleKana
                    seasonYear
                    seasonName
                    episodesCount
                    viewerStatusState
                }
                episode {
                    number
                    numberText
                }
                comment
                updatedAt
            }
        }
    }
}
'@

#LibraryEntryデータ取得用GraphQLクエリ(番組メモ取得)
$queryLibraries = @'
query ($user: String!) {
    user (username: $user) {
        libraryEntries {
            nodes {
                work {
                    annictId
                    title
                    titleKana
                    seasonYear
                    seasonName
                    episodesCount
                    viewerStatusState
                }
                note
            }
        }
    }
}
'@

# ImportExcel モジュールの存在を確認
function CheckImportExcelModule () {

    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "エラー: ImportExcel モジュールが見つかりません。"
        Write-Host "Excel出力をするには以下のコマンド(管理者権限)でインストールして下さい。"
        Write-Host "`nInstall-Module -Name ImportExcel`n"
        exit 1
    }
}


#アクセストークンチェック
function CheckAccessToken($accessToken) {

    if (-not $accessToken) {
        Write-Error "アクセストークン未設定: 下記URLで作成し"
        Write-Error "https://annict.com/settings/apps"
        Write-Error "Windowsのユーザー環境変数 ANNICT_ACCESS_TOKEN に設定してください"
        exit 1
    }
}

# GraphQLでAnnictからデータを取得
function AnnictApiRequest($query, $accessToken, $user) {

    $headers = @{ "Authorization" = "bearer $accessToken" }
    $body = if ($user) {
        @{ query = $query; variables = "{ ""user"": ""$user""}" } # $userがある場合GraphQL変数設定
    } else {
        @{ query = $query }
    }

    try {
        # GraphQLリクエスト送信
        $response = Invoke-RestMethod -Uri $endpoint -Method Post -Headers $headers -Body $body
        return $response
    } catch {
        Write-Error "Annict APIリクエストエラー: $_.Exception.Message"
        if ($_.Exception.Response) {
            Write-Error "HTTPステータスコード: $($_.Exception.Response.StatusCode)"
        } else {
            Write-Error "ネットワークエラーまたは認証失敗の可能性があります。"
        }
        return $null
    }
}

#デバッグ用JSONファイルの出力
function SaveDebugResponse($response) {

    $debugFile = "debug_response_$(get-date -Format 'yyyyMMdd_HHmmss_ffff').json"
    $debugFilePath = Join-Path -Path ([System.Environment]::GetFolderPath("Desktop")) -ChildPath $debugFile

    if(-not (Test-Path $debugFilePath)) {
        $response | ConvertTo-Json -Depth 10 | Out-File -FilePath $debugFilePath -Encoding UTF8
    }
    Write-Host "debug JSON File: $debugFilePath"
}

#データを取得しデータチェックをする
function RequestAndCheck($query, $accessToken, $dump, $user) {

    #データ取得
    $response = AnnictApiRequest $query $accessToken $user
    if (-not $response) {
        Write-Error "データ取得失敗 終了します"
        exit 1
    }

    #JSONレスポンス確認用
    if ($dump -eq $true) {
        SaveDebugResponse $response
    }

    #データチェック
    if (-not ($response.data)) {
        Write-Error "取得データエラー 終了します"
        SaveDebugResponse $response
        exit 1
    }
    return $response
}

#マッピング選択肢
enum Maps {
    STATUS
    RATING
    SEASON
    EPISODE
}

#マッピング用関数
function MapValue ( [Maps]$selectedMap, $value) {
    
    #選択された番号で今回のマッピングを設定
    switch ($selectedMap) {
        STATUS {
            # 視聴状況を置き換え Annictは見たい/見てる/見た/一時中断/視聴中止
            # Fani通は 繰り返し/視聴済/視聴途中/途中で切/初回切 なので適当に割当
            $mapping = @{
                "WATCHED"       = "視聴済"
                "WATCHING"      = "視聴途中"
                "WANNA_WATCH"   = "見たい"
                "ON_HOLD"       = "視聴途中"
                "STOP_WATCHING" = "途中で切"
            }
            break
        }
        RATING {
        # 評価を5-2に置き換え(Annictは4段階、Fani通は5段階なので5-2とする)
            $mapping = @{
                "GREAT"   = 5
                "GOOD"    = 4
                "AVERAGE" = 3
                "BAD"     = 2
            }
            break
        }
        SEASON {
            # シーズンのマッピング(ソート用に-\dを付加。出力時に削る)
            $mapping = @{
                "WINTER" = "-1冬"
                "SPRING" = "-2春"
                "SUMMER" = "-3夏"
                "AUTUMN" = "-4秋"
            }
            break
        }
        EPISODE {
            $mapping = @{
                0 = "単発"
            }
            break
        }
    }

    if ($null -eq $value) { return "" }

    #ContainsKeyが7系でInt64だと機能しないのでキャスト
    if ( 0 -eq $value) { [Int32]$value = $value }
    #マッピング
    $retValue = if ($mapping.ContainsKey($value)) { $mapping[$value] } else { $value }
    return $retValue
}

#カタカナをひらがなに変換
Add-Type -AssemblyName "Microsoft.VisualBasic"
function strConvHiragana($str) {
    return [Microsoft.VisualBasic.Strings]::StrConv($str, [Microsoft.VisualBasic.VbStrConv]::Hiragana)
}

# 取得したReviewJSONをCSVに変換
function ConvertAnnictRevewsToCsvdata($data) {

    #空のオブジェクトを持つitemを省く
    $items = $data.data.viewer.activities.edges | Where-Object { $_.item.PSObject.Properties.Count -gt 0 }

    #item一覧をFani通調査票CSVの1行に出力(項目がないものは空)
    $csvData = $items | ForEach-Object {
        [PSCustomObject]@{
            annictId    = $_.item.work.annictId
            作品タイトル = $_.item.work.title
            開始日       = "" # ([DateTime]::Parse($_.item.work.started_on)).ToLocalTime()
            終了日       = ""
            備考         = ""
            話数         = MapValue EPISODE $_.item.work.episodesCount
            視聴状況     = MapValue STATUS $_.item.work.viewerStatusState
            事前期待     = ""
            総合         = MapValue RATING $_.item.ratingOverallState
            初回         = ""
            ストーリー   = MapValue RATING $_.item.ratingStoryState
            ビジュアル   = MapValue RATING $_.item.ratingAnimationState
            キャスト     = MapValue RATING $_.item.ratingCharacterState
            楽曲         = MapValue RATING $_.item.ratingMusicState
            その他1項目名  = ""
            その他1項目評点 = ""
            その他2項目名   = ""
            その他2項目評点 = ""
            本文          = $_.item.body
            レビュー更新日 = ([DateTime]::Parse($_.item.updatedAt)).ToLocalTime()
            放映シーズン   = ($_.item.work.seasonYear).ToString() + (MapValue SEASON $_.item.work.seasonName)
            カナタイトル   = if($_.item.work.titleKana -eq "") { strConvHiragana $_.item.work.title} else {$_.item.work.titleKana}
#            カナタイトル  = $_.item.work.titleKana
        }
    }
    $sortedCsvData = $csvData | Sort-Object -Property シーズン, カナタイトル
    return $sortedCsvData
}

# 取得したRecordJSONをCSVに変換
function ConvertAnnictRecordsToCsvdata($data) {

    #コメントがnullのオブジェクトを省く
    $items = $data.data.user.records.nodes | Where-Object {$null -ne $_.comment}
    #Sortが7系でInt64だと機能しないのでキャスト
    $items = $items | Sort-Object -Property annictId, {[Int32]$_.episode.number}
    #item一覧をFani通調査票CSVの1行に出力(項目がないものは空)
    $csvData = $items | ForEach-Object {
        [PSCustomObject]@{
            annictId  = $_.work.annictId
            title     = $_.work.title
            titleKana = $_.work.titleKana
            season   = ($_.work.seasonYear).ToString() + (MapValue SEASON $_.work.seasonName)
            episodesCount = MapValue EPISODE $_.work.episodesCount
            viewerStatusState = MapValue STATUS $_.work.viewerStatusState
            body      = "$($_.episode.numberText):$($_.comment)"
            updatedAt = ([DateTime]::Parse($_.updatedAt)).ToLocalTime()
            merged    = $false
        }
    }
    return $csvData
}

# 取得したLibraryJSONをCSVに変換
function ConvertAnnictLibrariesToCsvdata($data) {

    #メモがnullのオブジェクトを省く
    $items = $data.data.user.libraryEntries.nodes | Where-Object {"" -ne $_.note}
    $items = $items | Sort-Object -Property annictId
    #item一覧をFani通調査票CSVの1行に出力(項目がないものは空)
    $csvData = $items | ForEach-Object {
        [PSCustomObject]@{
            annictId  = $_.work.annictId
            title     = $_.work.title
            titleKana = $_.work.titleKana
            season   = ($_.work.seasonYear).ToString() + (MapValue SEASON $_.work.seasonName)
            episodesCount = MapValue EPISODE $_.work.episodesCount
            viewerStatusState = MapValue STATUS $_.work.viewerStatusState
            body      = $_.note
            updatedAt = ""
            merged    = $false
        }
    }
    return $csvData
}

#Reviewデータに各話コメントもしくはメモをマージ
function MergeData ($reviews, $data) {

    foreach ($datum in $data) {
        foreach ($review in $reviews) {
            #annictIdが同一ならマージ
            if ($review.annictId -eq $datum.annictId) {
                $review.本文 += "`n$($datum.body)"
                $datum.merged = $true
                break
            }
        }
    }
    return $reviews
}

#マージ先レビューが無い各回コメントもしくはメモをレビューCSVに追加
function AddOphanData ($data, $orphans, $type) {

    foreach ($orphan in $orphans) {
        #マージ済はスキップ
        if ($orphan.Merged) {
            continue
        }

        $merged = $false
        if($type -eq "comment") {
            foreach ($datum in $data) {
                #1度追加されたレビューなし各回コメントがあれば追記
                if($datum.annictId -eq $orphan.annictId) {
                    $datum.本文 += "`n$($orphan.body)"
                    $datum.レビュー更新日 = $orphan.updatedAt
                    $merged = $true
                    break
                }
            }
        }

        #マージ先の無いコメントかメモを追加
        if(-not $merged) {
            $item = [PSCustomObject]@{
                annictId    = $orphan.annictId
                作品タイトル = $orphan.title
                開始日       = ""
                終了日       = ""
                備考         = if ($type -eq "comment") { "レビュー無し各話コメント" } else { "レビュー無し作品メモ" }
                話数         = $orphan.episodesCount
                視聴状況     = $orphan.viewerStatusState
                事前期待     = ""
                総合         = ""
                初回         = ""
                ストーリー   = ""
                ビジュアル   = ""
                キャスト     = ""
                楽曲         = ""
                その他1項目名  = ""
                その他1項目評点 = ""
                その他2項目名   = ""
                その他2項目評点 = ""
                本文          = $orphan.body
                レビュー更新日 = if ($type -eq "comment") { $orphan.updatedAt } else { "" }
                放映シーズン   = $orphan.season
                カナタイトル  = $orphan.titleKana
            }
            $data += $item    
        }
    }
    return $data
}

#CSVをソート
function SortCSV ($csv) {
    $csv = $csv | Sort-Object -Property 放映シーズン, カナタイトル
    #シーズンに付加してあるソート用の文字列-\dを削除
    $csv | ForEach-Object {$_.放映シーズン = $_.放映シーズン -replace "-\d", "" }
    return $csv
}

#CSV/Excelを出力
function ExportFile ($csv, $outputFilePath, $exportCsv) {

    if($exportCsv) {
        $outputFilePath += ".csv"
        try {
            # Excelで読み込めるutf8 CSV(BOM有)を指定する方法がバージョンで違う
            $psVersion = $PSVersionTable.PSVersion
            if ($psVersion.Major -gt 7 -or ($psVersion.Major -eq 7 -and $psVersion.Minor -ge 1)) {
                $csv | Export-Csv -Path $outputFilePath -NoTypeInformation -Encoding utf8BOM
            } else {
                $csv | Export-Csv -Path $outputFilePath -NoTypeInformation -Encoding utf8
            }
            Write-Host "CSV出力: $outputFilePath"
        } catch {
            Write-Error "CSV出力エラー: $($_.Exception.Message)"
            Write-Error "ファイルパス: $outputFilePath"
        }
    }
    else {
        $outputFilePath += ".xlsx"
        try {
            $csv | Export-Excel -Path $outputFilePath
            Write-Host "Excel出力: $outputFilePath"
        } catch {
            Write-Error "Excel出力エラー: $($_.Exception.Message)"
            Write-Error "ファイルパス: $outputFilePath"
        }
    }

}

#メイン処理
function Main($accessToken, $outputFilePath, $ags) {

    #引数-dumpがあるならTrue
    $dump = $ags -contains "-dump"
    #引数-csvがあるならTrue
    $exportCsv = $ags -contains "-csv"

    #ImportExcelモジュール存在チェック
    if(-not $exportCsv) {
        CheckImportExcelModule
    }

    #アクセストークンチェック
    CheckAccessToken $accessToken

    #Reviews取得(作品レビュー、評価)
    $response = RequestAndCheck $queryReviews $accessToken $dump ""
    #CSVデータに変換
    $csvdReviews = ConvertAnnictRevewsToCsvdata $response
    
    #利用者名取得
    $response = RequestAndCheck $queryName $accessToken $dump ""
    $username = $response.data.viewer.name

    #Records取得(各回コメント)
    $response = RequestAndCheck $queryRecords $accessToken $dump $username
    #CSVデータに変換
    $csvRecords = ConvertAnnictRecordsToCsvdata $response

    #Library取得(メモ(note))
    $response = RequestAndCheck $queryLibraries $accessToken $dump $username
    #CSVデータに変換
    $csvLibraries = ConvertAnnictLibrariesToCsvdata $response

    #Reviewsに各回コメント, メモをマージ
    $csvdata = MergeData $csvdReviews $csvLibraries
    $csvdata = MergeData $csvdata $csvRecords

    #マージ先なしメモ(note)、各回コメントを追加
    $csvdata = AddOphanData $csvdata $csvLibraries "note"
    $csvdata = AddOphanData $csvdata $csvRecords "comment"

    #CSVdataをソート
    $csvData = SortCSV $csvdata

    #CSV/Excel出力
    ExportFile $csvdata $outputFilePath $exportCsv

    Start-Sleep -Seconds 5
}

#メイン実行
Main $accessToken $outputFilePath $args
