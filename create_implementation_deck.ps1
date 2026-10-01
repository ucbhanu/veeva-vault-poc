$ErrorActionPreference = 'Stop'
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1
$presentation = $ppt.Presentations.Add()
$presentation.PageSetup.SlideWidth = 960
$presentation.PageSetup.SlideHeight = 540

function Color([int]$r, [int]$g, [int]$b) { return ($r -bor ($g -shl 8) -bor ($b -shl 16)) }
$ink = Color 35 43 52
$muted = Color 99 112 124
$teal = Color 24 126 122
$green = Color 91 151 113
$lime = Color 189 207 106
$pale = Color 245 247 246
$white = Color 255 255 255
$line = Color 220 227 225
$dark = Color 34 43 52

function Add-Rect($slide, [single]$x, [single]$y, [single]$w, [single]$h, [int]$fill, [int]$border = -1) {
    $shape = $slide.Shapes.AddShape(1, $x, $y, $w, $h)
    $shape.Fill.ForeColor.RGB = $fill
    if ($border -eq -1) { $border = $fill }
    $shape.Line.ForeColor.RGB = $border
    return $shape
}
function Add-Text($slide, [string]$text, [single]$x, [single]$y, [single]$w, [single]$h, [single]$size, [int]$color, [bool]$bold = $false) {
    $shape = $slide.Shapes.AddTextbox(1, $x, $y, $w, $h)
    $shape.TextFrame.TextRange.Text = $text
    $shape.TextFrame.TextRange.Font.Name = 'Aptos'
    $shape.TextFrame.TextRange.Font.Size = $size
    $shape.TextFrame.TextRange.Font.Color.RGB = $color
    $shape.TextFrame.TextRange.Font.Bold = $bold
    $shape.TextFrame.MarginLeft = 0
    $shape.TextFrame.MarginRight = 0
    $shape.TextFrame.MarginTop = 0
    $shape.TextFrame.MarginBottom = 0
    return $shape
}
function Add-BaseSlide([string]$title, [int]$number) {
    $slide = $presentation.Slides.Add($presentation.Slides.Count + 1, 12)
    $slide.Background.Fill.ForeColor.RGB = $pale
    Add-Rect $slide 0 0 960 8 $teal | Out-Null
    Add-Text $slide $title 52 32 850 42 26 $ink $true | Out-Null
    Add-Text $slide ("VEeva LITIGATION EXPORT  /  TWO-WEEK PLAN                                      {0:D2}" -f $number) 52 510 850 16 9 $muted $false | Out-Null
    return $slide
}

$slide = $presentation.Slides.Add(1, 12)
$slide.Background.Fill.ForeColor.RGB = $dark
Add-Rect $slide 0 0 18 540 $teal | Out-Null
Add-Text $slide 'VEEVA LITIGATION' 62 72 820 48 18 $lime $true | Out-Null
Add-Text $slide 'Export automation' 62 128 820 60 38 $white $true | Out-Null
Add-Text $slide 'Two-week implementation plan' 64 218 780 36 22 $lime $false | Out-Null
Add-Text $slide 'A controlled path from Vault request to validated, structured delivery.' 64 292 760 54 18 $white $false | Out-Null
Add-Text $slide '10 BUSINESS DAYS   |   STAGE-VALIDATED SOURCE DOWNLOAD   |   ZIP DELIVERY' 64 453 820 24 11 $lime $true | Out-Null
Add-Text $slide 'PLAN  /  01' 64 489 220 20 10 $white $false | Out-Null

$slide = Add-BaseSlide 'Executive summary' 2
$cards = @(
    @('PROOF IN PLACE', 'Stage authentication, metadata, filtered query, a real source download and structured ZIP have been exercised.'),
    @('TARGET STATE', 'Request-driven source, audit and annotation collection with validation, manifest, exceptions and controlled delivery.'),
    @('TWO-WEEK OUTCOME', 'Pilot-ready automation, approved access, operational runbook, UAT evidence and a go-live decision.')
)
$x = 52
foreach ($card in $cards) {
    Add-Rect $slide $x 112 268 210 $white $line | Out-Null
    Add-Rect $slide $x 112 268 5 $teal | Out-Null
    Add-Text $slide $card[0] ($x + 18) 137 230 24 12 $teal $true | Out-Null
    Add-Text $slide $card[1] ($x + 18) 180 230 118 16 $ink $false | Out-Null
    $x += 286
}
Add-Text $slide 'PILOT EVIDENCE' 54 357 170 22 11 $teal $true | Out-Null
Add-Text $slide 'One authorized source version was downloaded and packaged by document number. Wider searches exposed lifecycle permissions and unavailable content; those exceptions are explicit work items.' 54 386 830 68 15 $ink $false | Out-Null

$slide = Add-BaseSlide 'Target automated workflow' 3
$steps = @(
    @('01', 'Request intake', 'Request ID, criteria, naming rule, destination and approval'),
    @('02', 'Vault access', 'Environment, proxy trust, authentication and least privilege'),
    @('03', 'Find versions', 'Filtered VQL, major/minor versions, filename__v and paging'),
    @('04', 'Collect artifacts', 'Source content plus approved audit and per-version annotation routes'),
    @('05', 'Validate & package', 'Completeness, checksums, exceptions, folders, ZIP and manifest'),
    @('06', 'Deliver & audit', 'Approved SharePoint/Teams location, receipt, logs and retention')
)
$y = 98
foreach ($step in $steps) {
    Add-Rect $slide 56 $y 45 40 $teal | Out-Null
    Add-Text $slide $step[0] 62 ($y + 10) 34 20 11 $white $true | Out-Null
    Add-Text $slide $step[1] 119 ($y + 1) 196 23 15 $ink $true | Out-Null
    Add-Text $slide $step[2] 322 ($y + 2) 565 34 12 $muted $false | Out-Null
    $y += 62
}

$slide = Add-BaseSlide 'Two-week implementation Gantt' 4
Add-Text $slide 'WORKSTREAM' 54 99 295 20 10 $muted $true | Out-Null
$gridX = 350
$cell = 54
for ($day = 1; $day -le 10; $day++) {
    Add-Text $slide ("D{0}" -f $day) ($gridX + (($day - 1) * $cell)) 99 42 20 10 $muted $true | Out-Null
}
$tasks = @(
    @('Kickoff, scope, security & access', 1, 2, $teal),
    @('Requirements and process mapping', 1, 3, $green),
    @('Vault auth, session and environments', 2, 4, $teal),
    @('Version query and source retrieval', 3, 6, $teal),
    @('Audit-trail and annotation route', 4, 7, $green),
    @('Manifest, checksums, completeness', 5, 8, $lime),
    @('ZIP, naming and delivery integration', 6, 8, $lime),
    @('Retries, exceptions and monitoring', 7, 9, (Color 221 174 76)),
    @('UAT, fixes and acceptance', 8, 10, (Color 213 133 77)),
    @('Runbook, training and go-live gate', 9, 10, (Color 197 105 82))
)
$y = 128
foreach ($task in $tasks) {
    Add-Text $slide $task[0] 54 $y 286 19 10 $ink $false | Out-Null
    for ($day = 1; $day -le 10; $day++) { Add-Rect $slide ($gridX + (($day - 1) * $cell)) ($y + 1) 44 14 $line | Out-Null }
    for ($day = $task[1]; $day -le $task[2]; $day++) { Add-Rect $slide ($gridX + (($day - 1) * $cell)) ($y + 1) 44 14 $task[3] | Out-Null }
    $y += 31
}
Add-Text $slide 'D1-D5  FOUNDATION & BUILD     /     D6-D8  INTEGRATION & CONTROLS     /     D9-D10  UAT & HANDOFF' 54 452 850 28 10 $muted $true | Out-Null

$slide = Add-BaseSlide 'Deliverables and acceptance gates' 5
$deliverables = @(
    'Approved request criteria and environment configuration',
    'All requested eligible source major/minor versions',
    'Audit trail export and per-version annotation procedure',
    'Document-number folders, naming rules and clean ZIP',
    'Manifest with version, filename, ID, status and path',
    'Checksums, completeness check and actionable exceptions',
    'Delivery receipt, logs, runbook and support ownership'
)
$y = 106
foreach ($item in $deliverables) {
    Add-Text $slide '-' 60 $y 18 20 16 $teal $true | Out-Null
    Add-Text $slide $item 84 $y 405 25 12 $ink $false | Out-Null
    $y += 39
}
Add-Rect $slide 530 112 366 324 $white $line | Out-Null
Add-Text $slide 'ACCEPTANCE GATES' 552 133 320 24 12 $teal $true | Out-Null
Add-Text $slide "01   Access is approved and auditable`n`n02   Versions reconcile to the request`n`n03   Files open and hashes validate`n`n04   Exceptions are visible and actionable`n`n05   ZIP structure and delivery pass UAT" 552 176 320 244 12 $ink $false | Out-Null

$slide = Add-BaseSlide 'Risks, dependencies and decisions' 6
$risks = @(
    @('Vault permissions', 'Download Source varies by lifecycle/status.', 'Confirm roles and representative eligible records.'),
    @('Artifact pathways', 'Audit and annotation access may differ from source API.', 'Agree supported automation route by Day 3.'),
    @('Proxy and secrets', 'Corporate routing and CA trust are environment-specific.', 'Use approved secret store and trusted certificates.'),
    @('Scale and reliability', 'Large searches, throttling and missing content.', 'Page, retry, resume, checksum and report errors.'),
    @('Delivery governance', 'Customer destination, retention and approval.', 'Name destination owner and authorize release gate.')
)
$y = 105
foreach ($risk in $risks) {
    Add-Rect $slide 54 $y 850 64 $white $line | Out-Null
    Add-Text $slide $risk[0] 68 ($y + 10) 174 22 12 $teal $true | Out-Null
    Add-Text $slide $risk[1] 246 ($y + 8) 304 48 10 $ink $false | Out-Null
    Add-Text $slide $risk[2] 563 ($y + 8) 324 48 10 $muted $false | Out-Null
    $y += 73
}

$slide = Add-BaseSlide 'Operating model after go-live' 7
$roles = @(
    @('REQUESTER', "Submit approved request`nSet exact document criteria`nReview exception list"),
    @('AUTOMATION', "Authenticate and query`nRetrieve and validate`nPackage and record"),
    @('VAULT / IT', "Maintain roles and proxy trust`nSupport APIs and environment`nApprove retention controls"),
    @('SERVICE OWNER', "Monitor runs and failures`nReconcile manifest to request`nAuthorize and close delivery")
)
$x = 54
foreach ($role in $roles) {
    Add-Rect $slide $x 142 205 218 $white $line | Out-Null
    Add-Text $slide $role[0] ($x + 16) 162 174 24 12 $teal $true | Out-Null
    Add-Text $slide $role[1] ($x + 16) 207 174 125 12 $ink $false | Out-Null
    $x += 216
}
Add-Text $slide 'GO-LIVE DECISION  /  Permissions confirmed + representative UAT passed + delivery destination approved + service owner assigned' 55 400 842 48 13 $ink $true | Out-Null

$slide = Add-BaseSlide 'Solution structure' 8
Add-Rect $slide 54 102 382 330 $dark $dark | Out-Null
Add-Text $slide 'veeva-simple-poc/' 76 123 320 23 15 $lime $true | Out-Null
Add-Text $slide "|-- veeva_download.py     API retrieval`n|-- zip_creator.py        organize + ZIP`n|-- sharepoint_upload.py  delivery upload`n|-- .env                  local configuration`n|-- input/                downloaded artifacts`n|-- output/               delivery packages`n|-- README.md             operating guide`n'-- output/*.pptx         implementation plan" 76 162 325 220 13 $white $false | Out-Null
Add-Rect $slide 486 102 414 330 $white $line | Out-Null
Add-Text $slide 'Runtime flow' 512 124 320 25 15 $teal $true | Out-Null
Add-Text $slide "1  Stage auth -> returned Vault API host`n`n2  Metadata check -> product__v`n`n3  VQL -> allversions documents`n`n4  Source download -> document/version/Source`n`n5  Manifest -> filename, version, ID, status`n`n6  ZIP creator -> Source / Audit_Trail / Annotations" 512 163 345 235 13 $ink $false | Out-Null

$slide = Add-BaseSlide 'Code walkthrough' 9
Add-Rect $slide 54 103 410 330 $dark $dark | Out-Null
Add-Text $slide 'veeva_download.py' 76 124 340 24 15 $lime $true | Out-Null
Add-Text $slide "authenticate(auth_url, user, password)`n  returns sessionId + Vault API root`n`nquery_documents(api_root, version, session, search)`n  returns allversions + major/minor + filename__v`n`ndownload_document_version(...)`n  calls the version file endpoint`n  writes input/document/Source/versioned-file" 76 164 350 220 12 $white $false | Out-Null
Add-Rect $slide 490 103 410 330 $white $line | Out-Null
Add-Text $slide 'zip_creator.py' 514 124 340 24 15 $teal $true | Out-Null
Add-Text $slide "DOCUMENT_PATTERN`n  preserves IDs such as MA-AFL_8mg-DE-0297`n`ncategory_for(filename)`n  routes Source / Audit_Trail / Annotations`n`nprocess(input, request_id, output)`n  writes collection_tracker.csv`n  writes exceptions_report.csv`n  creates the request delivery ZIP" 514 164 350 220 12 $ink $false | Out-Null
Add-Text $slide 'Configuration: VEEVA_INPUT_FOLDER | VEEVA_REQUEST_ID | VEEVA_DOCUMENT_NAME_CONTAINS | STAGE_AUTH_URL | VEEVA_PROXY' 54 454 846 26 10 $muted $true | Out-Null

$slide = Add-BaseSlide 'Vault API details' 10
Add-Rect $slide 54 101 850 328 $dark $dark | Out-Null
Add-Text $slide 'AUTHENTICATION' 76 123 200 22 13 $lime $true | Out-Null
Add-Text $slide "POST https://sbbayer-uat.veevavault.com/api/v26.2/auth`nForm: username, password`nResponse: sessionId, vaultId, vaultIds[].url" 76 153 370 72 12 $white $false | Out-Null
Add-Text $slide 'METADATA' 500 123 180 22 13 $lime $true | Out-Null
Add-Text $slide "GET {api_root}/v26.2/metadata/vobjects/product__v`nHeader: Authorization: {sessionId}`nPurpose: validate session and Vault access" 500 153 350 72 12 $white $false | Out-Null
Add-Text $slide 'DOCUMENT QUERY' 76 255 220 22 13 $lime $true | Out-Null
Add-Text $slide "POST {api_root}/v26.2/query`nVQL: SELECT id, name__v, filename__v, document_number__v,`nmajor_version_number__v, minor_version_number__v`nFROM allversions documents WHERE name__v CONTAINS (...)" 76 285 390 90 12 $white $false | Out-Null
Add-Text $slide 'VERSION FILE' 500 255 180 22 13 $lime $true | Out-Null
Add-Text $slide "GET {api_root}/v26.2/objects/documents/{id}/`nversions/{major}/{minor}/file`nBinary response -> document/Source/`n{document}_v{major}.{minor}_{filename__v}" 500 285 360 90 12 $white $false | Out-Null
Add-Text $slide 'TLS uses Windows trusted ROOT/CA certificates; corporate proxy is configured through VEEVA_PROXY.' 54 455 850 24 10 $muted $true | Out-Null

$slide = Add-BaseSlide 'Validation snapshots' 11
Add-Text $slide 'AUTH + QUERY' 54 96 250 20 12 $teal $true | Out-Null
Add-Rect $slide 54 122 410 126 $dark $dark | Out-Null
Add-Text $slide "Authenticated. vaultId=14605`nMetadata OK for object: product__v`nQuery matched 431 document version(s).`nfilename__v: 9df52c..._Test_2.txt" 72 143 370 90 12 $white $false | Out-Null
Add-Text $slide 'DOWNLOAD RESULT' 500 96 260 20 12 $teal $true | Out-Null
Add-Rect $slide 500 122 404 126 $white $line | Out-Null
Add-Text $slide "Document: MA-AFL_8mg-DE-0297`nVersion: 0.1`nBytes: 17,772`nStatus: DOWNLOADED" 520 143 350 90 13 $ink $false | Out-Null
Add-Text $slide 'INPUT STRUCTURE' 54 282 250 20 12 $teal $true | Out-Null
Add-Rect $slide 54 308 410 126 $dark $dark | Out-Null
Add-Text $slide "input/`n|-- MA-AFL_8mg-DE-0297/`n    |-- Source/`n        '-- ..._v0.1_...Test_2.txt" 72 328 370 90 12 $white $false | Out-Null
Add-Text $slide 'ZIP CONTENTS' 500 282 250 20 12 $teal $true | Out-Null
Add-Rect $slide 500 308 404 126 $white $line | Out-Null
Add-Text $slide "ENV-001-RESUMED_delivery.zip`n- collection_tracker.csv`n- exceptions_report.csv`n- document/Source/versioned-file" 520 328 350 90 12 $ink $false | Out-Null
Add-Text $slide 'Snapshots summarize the verified local run; credentials and session tokens are intentionally omitted.' 54 455 850 24 10 $muted $true | Out-Null

$outputPath = Join-Path (Get-Location) 'output\Veeva_Litigation_Automation_2_Week_Plan.pptx'
$presentation.SaveAs($outputPath)
$presentation.Close()
$ppt.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
Write-Output "Created $outputPath"
Get-Item $outputPath | Select-Object Name, Length, LastWriteTime
