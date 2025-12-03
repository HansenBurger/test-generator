# PowerShell辅助脚本 - 用于文档解析和XMind生成
param(
    [Parameter(Mandatory=$true)]
    [string]$FilePath,
    
    [Parameter(Mandatory=$true)]
    [string]$ApiUrl
)

$ErrorActionPreference = "Stop"

try {
    # 1. 解析文档
    Write-Host "正在上传并解析文档..."
    $parseUri = "$ApiUrl/api/parse-doc"
    
    # 尝试使用-Form参数（PowerShell 6.0+），如果不支持则使用HttpClient
    $fileItem = Get-Item -Path $FilePath
    $psVersion = $PSVersionTable.PSVersion.Major
    
    if ($psVersion -ge 6) {
        # PowerShell 6.0+ 支持 -Form 参数
        $form = @{
            file = $fileItem
        }
        $parseResponse = Invoke-RestMethod -Uri $parseUri -Method Post -Form $form
    } else {
        # PowerShell 5.1 使用 HttpClient
        Add-Type -AssemblyName System.Net.Http
        
        $fileBytes = [System.IO.File]::ReadAllBytes($FilePath)
        $fileName = $fileItem.Name
        $boundary = [System.Guid]::NewGuid().ToString()
        
        # 构建multipart/form-data
        $bodyBuilder = New-Object System.Text.StringBuilder
        $bodyBuilder.Append("--$boundary`r`n") | Out-Null
        $bodyBuilder.Append("Content-Disposition: form-data; name=`"file`"; filename=`"$fileName`"`r`n") | Out-Null
        $bodyBuilder.Append("Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document`r`n`r`n") | Out-Null
        
        $headerBytes = [System.Text.Encoding]::UTF8.GetBytes($bodyBuilder.ToString())
        $footerBytes = [System.Text.Encoding]::UTF8.GetBytes("`r`n--$boundary--`r`n")
        
        $bodyList = New-Object System.Collections.Generic.List[byte]
        $bodyList.AddRange($headerBytes)
        $bodyList.AddRange($fileBytes)
        $bodyList.AddRange($footerBytes)
        
        $httpClient = New-Object System.Net.Http.HttpClient
        $content = New-Object System.Net.Http.ByteArrayContent($bodyList.ToArray())
        $content.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse("multipart/form-data; boundary=$boundary")
        
        try {
            $response = $httpClient.PostAsync($parseUri, $content).Result
            $responseContent = $response.Content.ReadAsStringAsync().Result
            
            if (-not $response.IsSuccessStatusCode) {
                Write-Host "HTTP错误: $($response.StatusCode) - $responseContent" -ForegroundColor Red
                exit 1
            }
            
            $parseResponse = $responseContent | ConvertFrom-Json
        } finally {
            $httpClient.Dispose()
            $content.Dispose()
        }
    }
    
    if (-not $parseResponse.success) {
        Write-Host "解析失败: $($parseResponse.message)" -ForegroundColor Red
        exit 1
    }
    
    Write-Host "文档解析成功！" -ForegroundColor Green
    
    # 2. 生成XMind文件
    Write-Host "正在生成XMind测试大纲..."
    $generateUri = "$ApiUrl/api/generate-outline-from-json"
    
    $jsonBody = $parseResponse.data | ConvertTo-Json -Depth 10 -Compress
    $generateResponse = Invoke-WebRequest -Uri $generateUri -Method Post -Body $jsonBody -ContentType "application/json; charset=utf-8"
    
    # 3. 保存文件
    $caseName = $parseResponse.data.requirement_info.case_name
    $version = $parseResponse.data.version
    
    if ($version) {
        $outputFileName = "$caseName-$version.xmind"
    } else {
        $outputFileName = "$caseName.xmind"
    }
    
    # 清理文件名中的非法字符
    $outputFileName = $outputFileName -replace '[\\/:*?"<>|]', '_'
    
    $outputDir = Split-Path -Path $FilePath -Parent
    $outputPath = Join-Path -Path $outputDir -ChildPath $outputFileName
    
    [System.IO.File]::WriteAllBytes($outputPath, $generateResponse.Content)
    
    Write-Host "XMind文件已生成: $outputPath" -ForegroundColor Green
    return $outputPath
    
} catch {
    Write-Host "错误: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.Exception.Response) {
        $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
        $responseBody = $reader.ReadToEnd()
        Write-Host "响应详情: $responseBody" -ForegroundColor Yellow
    }
    exit 1
}

