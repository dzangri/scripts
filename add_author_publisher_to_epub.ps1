function Add-EpubAuthorAndPublisher {
    param (
        [string]$epubFilePath,
        [string]$authorName,
        [string]$publisherName
    )

    # Check if the EPUB file exists
    if (-Not (Test-Path $epubFilePath -PathType Leaf)) {
        Write-Host "Error: The specified EPUB file '$epubFilePath' does not exist."
        return
    }

    try {
        # Create a temporary directory to extract the EPUB contents
        $tempDir = Join-Path -Path $env:TEMP -ChildPath ([System.Guid]::NewGuid().ToString())
        New-Item -ItemType Directory -Path $tempDir | Out-Null

        # Rename to zip since because Expand-Archive only accepts files with '.zip' explicitly
        $epubFile = Get-Item $epubFilePath
        $epubFileZip = $epubFile.Basename + '.zip'
        Rename-Item -Path $epubFilePath -NewName $epubFileZip

        # Extract the EPUB contents to the temporary directory
        Expand-Archive -Path $epubFileZip -DestinationPath $tempDir

        # Find the metadata.opf file within the EPUB package
        $metadataPath = Get-ChildItem -Path $tempDir -Recurse -Filter "metadata.opf" | Select-Object -First 1
        if (-not $metadataPath) {
            Write-Host "Error: Failed to find the 'metadata.opf' file within the EPUB package."
            return
        }

        # Load the metadata file as an XML document with a custom namespace manager
        $metadataXml = New-Object Xml.XmlDocument
        $metadataXml.Load($metadataPath.FullName)
        $xmlNamespaceManager = New-Object System.Xml.XmlNamespaceManager($metadataXml.NameTable)
        $xmlNamespaceManager.AddNamespace("dc", "http://purl.org/dc/elements/1.1/")
        # Infuriatingly, many .xml files have a generic namespace that looks different than others, and it needs to
        # be explicitly added to the namespace manager
        $xmlNamespaceManager.AddNamespace("ns", $metadataXml.DocumentElement.NamespaceURI)

        # Find or create the <dc:creator> element in the metadata
        $creatorNode = $metadataXml.SelectSingleNode('//dc:creator', $xmlNamespaceManager)
        if (-not $creatorNode) {
            # Create the <dc:creator> element and set its value
            $creatorNode = $metadataXml.CreateElement('dc', 'creator', $metadataXml.DocumentElement.NamespaceURI)
            $creatorNode.InnerText = $authorName

            # Append the <dc:creator> element to the metadata
            $metadataNode = $metadataXml.SelectSingleNode("//ns:metadata", $xmlNamespaceManager)
            $metadataNode.AppendChild($creatorNode)
        } else {
            # Update the existing <dc:creator> element with the new author name
            $creatorNode.InnerText = $authorName
        }

        # Find or create the <dc:publisher> element in the metadata
        $publisherNode = $metadataXml.SelectSingleNode('//dc:publisher', $xmlNamespaceManager)
        if (-not $publisherNode) {
            # Create the <dc:publisher> element and set its value
            $publisherNode = $metadataXml.CreateElement('dc', 'publisher', $metadataXml.DocumentElement.NamespaceURI)
            $publisherNode.InnerText = $publisherName

            # Append the <dc:publisher> element to the metadata
            $metadataNode = $metadataXml.SelectSingleNode("//ns:metadata", $xmlNamespaceManager)
            $metadataNode.AppendChild($publisherNode)
        } else {
            # Update the existing <dc:publisher> element with the new author name
            $publisherNode.InnerText = $publisherName
        }

        # Find or create the <dc:title> element in the metadata
        $titleNode = $metadataXml.SelectSingleNode('//dc:title', $xmlNamespaceManager)
        if (-not $titleNode) {
            Write-Host "Error: No existing title found in the .epub metadata."
            return
        }

        # Update the title by removing underscores and dashes and replacing them with spaces
        $title = $titleNode.InnerText -replace '[_-]', ' '
        $titleNode.InnerText = $title

        # Save the updated metadata back to the file
        $metadataXml.Save($metadataPath.FullName)

        # Re-create the EPUB file with the updated metadata in the temp directory
        $updatedEpubFilePath = Join-Path -Path $env:TEMP -ChildPath ([System.Guid]::NewGuid().ToString() + ".zip")
        Compress-Archive -Path $tempDir\* -DestinationPath $updatedEpubFilePath -Force

        #  Move the file back to the current one
        Move-Item -Path $updatedEpubFilePath -Destination $epubFilePath

        Write-Host "Author '$authorName' added to the EPUB file '$epubFilePath'."

        # Clean up the temporary directory
        Remove-Item -Path $tempDir -Recurse -Force

        # Remove zip of old file
        Remove-Item -Path $epubFileZip
    }
    catch {
        Write-Host "Error: $_.Exception.Message"
    }
}

# Example usage
$epubFilePath = $args[0]
$authorName = $args[1]
$publisherName = $args[2]

Add-EpubAuthorAndPublisher -epubFilePath $epubFilePath -authorName $authorName -publisherName $publisherName
