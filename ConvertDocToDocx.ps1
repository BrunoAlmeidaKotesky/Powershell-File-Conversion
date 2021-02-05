function Convert-WordDocument
{
  param
  (
    # accept path strings or items from Get-ChildItem
    [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName)]
    [string]
    [Alias('FullName')]
    $Path
  )
  
  begin
  {
    # we are collecting all paths first
    [Collections.ArrayList]$collector = @()
  }

  process
  {
    # find extension
    $extension = [System.IO.Path]::GetExtension($Path)
    
    # we only process .doc and .dot files
    if ($extension -eq '.doc' -or $extension -eq '.dot')
    {
        # add to list for later processing
        $null = $collector.Add($Path)

    }
  }
  end
  {   
    # pipeline is done, now we can start converting!

    Write-Progress -Activity Converting -Status 'Launching Application'

    # initialize Word (must be installed)
    $word = New-Object -ComObject Word.Application

    $counter = 0
    Foreach ($Path in $collector)
    {
        # increment a counter for the progress bar
        $counter++

        # open document in Word
        $doc = $word.Documents.Open($Path)

        # determine target document type
        # if the doc has macros, use different extensions

        [string]$targetExtension = ''
        [int]$targetConversion = 0

        switch ([System.IO.Path]::GetExtension($Path))
        { 
          '.doc' {    
            if ($doc.HasVBProject -eq $true)
            { 
              $targetExtension = '.docm'
              $targetConversion = 13
            }
            else
            {
              $targetExtension = '.docx'  
              $targetConversion = 16     
            }
          }
          '.dot' {
            if ($doc.HasVBProject -eq $true)
            { 
              $targetExtension = '.dotm'
              $targetConversion = 15 
            }
            else
            {
              $targetExtension = '.dotx'  
              $targetConversion = 14      
            }
          }
        }

        # conversion cannot work for read-only docs
        If (!$doc.ActiveWindow.View.ReadingLayout)
        {
            if ($targetConversion -gt 0)
            {
              $pathOut = [IO.Path]::ChangeExtension($Path, $targetExtension)
              
              $doc.Convert()
              $percent = $counter * 100 / $collector.Count
              Write-Progress -Activity 'Converting' -Status $pathOut -PercentComplete $percent
              $doc.SaveAs([ref]$PathOut,[ref] $targetConversion)
            }
        }

        $word.ActiveDocument.Close()
    } 

    # quit Word when done
    Write-Progress -Activity Converting -Status Done.
    $word.Quit()
  }
}

Convert-WordDocument -Path "C:\Docs\Doc.doc";