[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$InputPptx,
    [Parameter(Mandatory = $true)][string]$OutJson,
    [Parameter(Mandatory = $true)][string]$ShapeIds
)

# Reports how PowerPoint itself breaks a text box into display lines, so a long
# abstract can be split across the template's two boxes without guessing.
# The .pptx is opened from a scratch copy and never saved.
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Release-ComObject([object]$Object) {
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
    }
}

$source = (Resolve-Path -LiteralPath $InputPptx).ProviderPath
$jsonPath = [IO.Path]::GetFullPath($OutJson)
$wanted = @($ShapeIds -split ',' | ForEach-Object { [int]$_.Trim() })

$scratch = Join-Path ([IO.Path]::GetTempPath()) ('.measure-' + [Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($scratch)
$copy = Join-Path $scratch 'input.pptx'
Copy-Item -LiteralPath $source -Destination $copy

$app = $null
$presentations = $null
$deck = $null
$slides = $null
$slide = $null
$shapes = $null
$result = [Collections.Generic.List[object]]::new()

try {
    $app = New-Object -ComObject PowerPoint.Application
    $presentations = $app.Presentations
    $deck = $presentations.Open($copy, -1, 0, 0)
    $slides = $deck.Slides
    $slide = $slides.Item(1)
    $shapes = $slide.Shapes
    foreach ($id in $wanted) {
        $shape = $null
        $frame = $null
        $range = $null
        $lines = $null
        try {
            $shape = $null
            for ($i = 1; $i -le $shapes.Count; $i++) {
                if ([int]$shapes.Item($i).Id -eq $id) { $shape = $shapes.Item($i); break }
            }
            if ($null -eq $shape) { throw "shape $id not found" }
            $frame = $shape.TextFrame2
            $range = $frame.TextRange
            $lines = $range.Lines()
            $count = [int]$lines.Count
            $infos = [Collections.Generic.List[object]]::new()
            for ($j = 1; $j -le $count; $j++) {
                $one = $null
                try {
                    $one = $range.Lines($j, 1)
                    $infos.Add([ordered]@{
                        start = [int]$one.Start
                        text  = ([string]$one.Text).TrimEnd([char]13)
                    })
                } finally { Release-ComObject $one }
            }
            $result.Add([ordered]@{
                id        = $id
                name      = [string]$shape.Name
                text      = [string]$range.Text
                chars     = ([string]$range.Text).Length
                line_count = $count
                box_pt    = [ordered]@{ left = [double]$shape.Left; top = [double]$shape.Top; width = [double]$shape.Width; height = [double]$shape.Height }
                lines     = @($infos.ToArray())
            })
        } finally {
            Release-ComObject $lines
            Release-ComObject $range
            Release-ComObject $frame
            Release-ComObject $shape
        }
    }
    $json = [ordered]@{ source = $source; shapes = @($result.ToArray()) }
    $json | ConvertTo-Json -Depth 8 | Out-File -LiteralPath $jsonPath -Encoding utf8
    [ordered]@{ json = $jsonPath; shapes = @($result | ForEach-Object { [ordered]@{ id = $_.id; lines = $_.line_count; chars = $_.chars } }) } | ConvertTo-Json -Depth 4
} finally {
    Release-ComObject $shapes
    Release-ComObject $slide
    Release-ComObject $slides
    if ($null -ne $deck) { try { $deck.Close() } finally { Release-ComObject $deck } }
    Release-ComObject $presentations
    Release-ComObject $app
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if (Test-Path -LiteralPath $scratch) { Remove-Item -LiteralPath $scratch -Recurse -Force }
}
