Param(
    # Filename of the document
    [Parameter(Mandatory=$true)]
    [string] $File,
    # Text to be replaced
    [Parameter(Mandatory=$true)]
    [string]
    $FindText,
    # Text to replace with
    [Parameter(Mandatory=$true)]
    [string]
    $ReplaceWith
)

Add-Type -AssemblyName Microsoft.Office.Interop.Word

$objWord = New-Object -ComObject Word.Application  
$objDoc = $objWord.Documents.Open($File)
$objDoc.TrackFormatting = $false;

function Replace(
    [Parameter(Mandatory=$true)]
    $find
)
{
    $findWrap = [Microsoft.Office.Interop.Word.WdFindWrap]::wdFindContinue
    $replace = [Microsoft.Office.Interop.Word.WdReplace]::wdReplaceAll

    $find.Execute($FindText, 
                        $false, #match case
                        $false, #match whole word
                        $false, #match wildcards
                        $false, #match soundslike
                        $false, #match all word forms
                        $true,  #forward
                        $findWrap, 
                        $null,  #format
                        $ReplaceWith,
                        $replace)
}

function Rename-TextFrames()
{
    foreach($shape in $objDoc.Shapes)
    {
        $find = $shape.TextFrame.ContainingRange.Find
        Replace -find $find
    }
}

function Rename-DocumentItems()
{
    foreach($section in $objDoc.Sections)
    {
        foreach($header in $section.Headers)
        {
            $find = $header.Range.Find
            Replace -find $find
            $header.Shapes | ForEach-Object {
                if ($_. Type -eq [Microsoft.Office.Core.MsoShapeType])
                {
                    $find = $_.TextFrame.TextRange.Find
                    Replace -find $find
                }
            }
        }
        foreach ($footer in $section.Footers)
        {
            $find = $footer.Range.Find
            Replace -find $find
        }
    }
    Rename-TextFrames
}

Rename-DocumentItems

$objDoc.Save()
$objDoc.Close()
$objWord.Quit()

Write-Host("Finished document: " + $File);
