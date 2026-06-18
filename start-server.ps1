# Zoho Outlook Add-in — local dev server
# Run this script first, then sideload manifest.xml in Outlook.
# Keep the window open while using the add-in.

$root = $PSScriptRoot

# --- Generate icon PNGs using System.Drawing ---
function New-IconPng($path, $size) {
    Add-Type -AssemblyName System.Drawing
    $bmp = New-Object System.Drawing.Bitmap($size, $size)
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $g.Clear([System.Drawing.Color]::FromArgb(228, 75, 34))   # Zoho orange

    # Draw a simple "Z" letter
    $font  = New-Object System.Drawing.Font("Arial", [Math]::Max(6, $size * 0.55), [System.Drawing.FontStyle]::Bold)
    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
    $fmt   = New-Object System.Drawing.StringFormat
    $fmt.Alignment     = [System.Drawing.StringAlignment]::Center
    $fmt.LineAlignment = [System.Drawing.StringAlignment]::Center
    $rect  = New-Object System.Drawing.RectangleF(0, 0, $size, $size)
    $g.DrawString("Z", $font, $brush, $rect, $fmt)

    $g.Dispose()
    $dir = Split-Path $path
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
}

$assetsDir = Join-Path $root "assets"
if (-not (Test-Path (Join-Path $assetsDir "icon-16.png"))) {
    Write-Host "Generating icon assets..." -ForegroundColor Cyan
    New-IconPng (Join-Path $assetsDir "icon-16.png") 16
    New-IconPng (Join-Path $assetsDir "icon-32.png") 32
    New-IconPng (Join-Path $assetsDir "icon-80.png") 80
    Write-Host "Icons created." -ForegroundColor Green
}

# --- MIME types ---
$mime = @{
    ".html" = "text/html; charset=utf-8"
    ".css"  = "text/css"
    ".js"   = "application/javascript"
    ".png"  = "image/png"
    ".ico"  = "image/x-icon"
    ".json" = "application/json"
    ".xml"  = "application/xml"
}

# --- HTTP server ---
$port     = 3000
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")
$listener.Start()

Write-Host ""
Write-Host "Add-in server running at http://localhost:$port" -ForegroundColor Green
Write-Host "Press Ctrl+C to stop." -ForegroundColor Gray
Write-Host ""

try {
    while ($listener.IsListening) {
        $ctx = $listener.GetContext()
        $req = $ctx.Request
        $res = $ctx.Response

        $urlPath = $req.Url.AbsolutePath.TrimStart("/")
        if ($urlPath -eq "") { $urlPath = "taskpane.html" }

        $filePath = Join-Path $root $urlPath

        if (Test-Path $filePath -PathType Leaf) {
            $ext         = [System.IO.Path]::GetExtension($filePath).ToLower()
            $contentType = if ($mime.ContainsKey($ext)) { $mime[$ext] } else { "application/octet-stream" }
            $bytes       = [System.IO.File]::ReadAllBytes($filePath)

            $res.ContentType   = $contentType
            $res.ContentLength64 = $bytes.Length
            $res.StatusCode    = 200
            $res.OutputStream.Write($bytes, 0, $bytes.Length)
        } else {
            $res.StatusCode = 404
            $body = [System.Text.Encoding]::UTF8.GetBytes("404 Not Found: $urlPath")
            $res.ContentLength64 = $body.Length
            $res.OutputStream.Write($body, 0, $body.Length)
        }

        $res.OutputStream.Close()
    }
} finally {
    $listener.Stop()
}
