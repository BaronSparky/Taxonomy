#############################################################################################################
# 
# Name:
#       MergeXMLs.PS1
#
# Description:
#       This PowerShell script can be used to take the RDM XSL transformed file, an export from the current
#       Live termset from 02-ExportTermSet.ps1, and mix them together to form an new updated XML file that
#       can then be used to re-import into the MMS via 03-ImportNew-Taxonomy.ps1. This updates the Live
#       termset with all of the extra RDM output data such as their labels and the CVRObjectIDs which is a
#       custom property from the current RDM tool. This ObjectID is needed for later processes when we need
#       to manage change from RDM and reflect that in the MMS. 
#
#       NOTE: This doesn't help to manage change, only for enrichment of current termset with RDM data
#
# Author: 
#       Nigel Bridport, Microsoft Consulting Services
# Date: 
#       28/04/2016
# Version: 
#       V1.0
#
# Changes:
#       Version Date Who Description
#
#
# Inputs:
#       -SRCExport - The transformed file from the termset  
#       -LIVEExport - The Live export for the termset
#
# Example:
#       -SRCExport = "C:\Users\nigelbri\PS Scripts\Taxonomy Management\aaAFRICA_RDM.xml"
#       -LIVEExport = "C:\Users\nigelbri\PS Scripts\Taxonomy Management\aaAFRICA_LIVE.XML"
#
#############################################################################################################
#
Param(
    [Parameter(Mandatory=$false)]
    [string]$FilePath,
    
    [Parameter(Mandatory=$false)]
    [string]$SRCExport,
    
    [Parameter(Mandatory=$false)]
    [string]$LIVEExport
)

#Some global counters and switches
$global:glTermsModified = 0      #Keep a count of modified terms for summary results
$global:glTermsAdded = 0         #Keep a count of the terms added for summary results
$global:tsBusinessOwner = $false #So we can do some specific actions to Business Owner
$global:glTermsCount = 1         #Current term for a progress counter
$global:glTermsApprox = 0        #Total terms in the Live export...for a progress counter


################################################################
#
# Cleanse-File 
# This function removes any odd characters from the supplied
# file as SPO MMS does not support some characters
#
################################################################
function Cleanse-File($sFilePath){

    (Get-Content $sFilePath) | ForEach-Object {
        $_ -replace "  ", " " -replace '“', '[' -replace '”', ']' -replace "–", "-" -replace "’", "'" -replace "＆", "&"
        } | Set-Content $sFilePath -Encoding Unicode

    return $sFilePath
}


################################################################
#
# Get-SeeAlso
# Get the terms "See Also" propert
#
################################################################
function Get-SeeAlso($xmlLocalCustomProps){
    #We have issues with "See Also". In XSLT, we cannot combine strings on the fly, so the SRC input may have a
    # number of lines for "See Also" such as "See Also"/"See Also1"/"See Also2"/.../"See Also<n>"
    # We need to combine all of the "See Also" entries into a single value, ";" delimited. You cannot have multiple
    # properties with the same name. Se could just have multiple properties for "See Also" but it is neater to have them
    # all in a single property
    $sSeeAlsoValues = "" #Use this as a temporary holding string

    foreach ($xmlLocalCustomPropSRC in $xmlLocalCustomPropsSRC.LocalCustomProperty)
    {
        if ($xmlLocalCustomPropSRC.Key.ToUpper() -ge "SEE ALSO") #This finds any key beginning with "See Also"
        {
            $sSeeAlsoValues += $xmlLocalCustomPropSRC.Value + ";"
        }
    }
    return $sSeeAlsoValues.Trim(";")
}


################################################################
#
# Unwanted-Terms
# Request from the user, what action is required for terms that
# shouldn't really be in the taxonomy
#
################################################################
function Unwanted-Terms($xmlTermInvalidNode, $xmlNonProcessedObjects, $mbTitle, $mbMessage, $sInvalidType){

    foreach ($xmlNonProcessedObject in $xmlNonProcessedObjects)
    {
        $sTermPath = Term-Path($xmlNonProcessedObject)
        Write-Host "  $sInvalidType - [" -NoNewline -ForegroundColor Red
        Write-Host $sTermPath -NoNewline -ForegroundColor Red
        Write-Host "] " -ForegroundColor Red
    }

    #Ask the user what to do with this type of invalid term
    $Deprecate = New-Object System.Management.Automation.Host.ChoiceDescription "&Deprecate", "Deprecate all of the Terms."
    $Remove = New-Object System.Management.Automation.Host.ChoiceDescription "&Remove", "Delete all of the Terms."
    $Nothing = New-Object System.Management.Automation.Host.ChoiceDescription "&Nothing", "Do nothing with the Terms."
    $mbOptions = [System.Management.Automation.Host.ChoiceDescription[]]($Deprecate, $Remove, $Nothing)

    $mbResult = $Host.UI.PromptForChoice($mbTitle, $mbMessage, $mbOptions, 0) 

    if (($mbResult -eq 0) -or ($mbResult -eq 1))
    {
        foreach ($xmlNonProcessedObject in $xmlNonProcessedObjects)
        {
            $xmlTermFIXUP = $xmlTermInvalidNode.OwnerDocument.CreateElement("Fixup")
            $xmlTermFIXUP.setAttribute("Name", $xmlNonProcessedObject.Name)
            $xmlTermFIXUP.setAttribute("Id", $xmlNonProcessedObject.Id)
            $xmlTermFIXUP.setAttribute("Processed", "Merged")

            if ($mbResult -eq 0)
            {
                $xmlTermFIXUP.setAttribute("Action", "Deprecate")
            }
            elseif ($mbResult -eq 1)
            {
                $xmlTermFIXUP.setAttribute("Action", "Remove")
            }

            $xmlTermFIXUP = $xmlTermInvalidNode.AppendChild($xmlTermFIXUP)
        }
    }
}


################################################################
#
# Check-Processed
# Check the XML file to ensure all TERMS have been assessed
#
################################################################
function Check-Processed($xmlDocument, $sXMLFile){

    $bInvalidTerms = $false

    #Create the XML Node in preparation
    $xmlTermInvalidNode = $xmlDocument.TermStores.TermStore.Groups.Group.TermSets.TermSet.OwnerDocument.CreateElement("InvalidTerms")
    
    #Check for duplicates
    $xmlNonProcessedObjects = $xmlDocument.SelectNodes("//Term[@Processed=""Merged-Duplicate""]")

    if ($xmlNonProcessedObjects.Count -gt 0)
    {
        Write-Host $xmlNonProcessedObjects.Count"duplicate terms have been found. These exist in the LIVE export but not the SOURCE export." -ForegroundColor Red -BackgroundColor Yellow
        $mbMessage = "What do you want to do with the duplicate Terms?  (These were found in the LIVE export but not the SOURCE export)"

        Unwanted-Terms $xmlTermInvalidNode $xmlNonProcessedObjects "Duplicate Terms" $mbMessage "Duplicate"
        $bInvalidTerms = $true
    }

    #Check for other issues
    $xmlNonProcessedObjects = $xmlDocument.SelectNodes("//Term[@Processed=""False""]")

    if ($xmlNonProcessedObjects.Count -gt 0)
    {
        Write-Host $xmlNonProcessedObjects.Count"terms have NOT been processed. These exist in the LIVE export but not the SOURCE export." -ForegroundColor Red -BackgroundColor Yellow
        $mbMessage = "What do you want to do with the unprocessed Terms?  (These were found in the LIVE export but not the SOURCE export)"

        Unwanted-Terms $xmlTermInvalidNode $xmlNonProcessedObjects "Unprocessed" $mbMessage "Unprocessed"
        $bInvalidTerms = $true
    }

    #Close the node off
    $xmlTermInvalidNode = $xmlDocument.TermStores.TermStore.Groups.Group.TermSets.TermSet.AppendChild($xmlTermInvalidNode)

    if (-not $bInvalidTerms)
    {
        Write-Host "The processed file [$sXMLFile] has had all terms processed and no issues were found." -ForegroundColor Green
    }
}


################################################################
#
# GetTerm-Count
# This function returns the number of terms in the fixup file
#
################################################################
function GetTerm-Count($xmlDocument){

    #"Processed" only has four valid options:
    # "No Changes" - There is no difference between the Live and Src
    # "Merged" - There is a difference between Live and Src and that term needs updating
    # "Merged-Duplicate" - There is a difference between Live and Src but this is a duplicate term that should be deleted
    # "False" - The term has not been processed as it hasn't met any of the scan criteria. Needs further investigation
    $xmlTerms = $xmlDocument.SelectNodes("//Term[@Processed=""No Changes""]")
    $lTotalTerms += $xmlTerms.Count

    $xmlTerms = $xmlDocument.SelectNodes("//Term[@Processed=""Merged""]")
    $lTotalTerms += $xmlTerms.Count

    $xmlTerms = $xmlDocument.SelectNodes("//Term[@Processed=""Merged-Duplicate""]")
    $lTotalTerms += $xmlTerms.Count

    $xmlTerms = $xmlDocument.SelectNodes("//Term[@Processed=""False""]")
    $lTotalTerms += $xmlTerms.Count

    return $lTotalTerms.ToString()
}


################################################################
#
# Compare-Terms
# This function works through the TERM nodes
#
################################################################
function Compare-Terms($xmlTermsLIVE, $xmlTermLIVE, $xmlTermSRC, $sMessageOut) {

    try
    {
        #Update the properties on the LIVE terms from the SRC master Term nodes
        # LIVE carries Id/IsDeprecated/IsAvailableForTagging/IsKeyword/IsReused/IsRoot/IsSourceTerm/CustomSortOrder/CustomProperties/LocalCustomProperties
        # SRC carries CVRObjectID (which also goes into LocalCustomProperties)/Labels

        #Term Display Name checks
        if ($xmlTermLIVE.Name.CompareTo($xmlTermSRC.Name) -ne 0)
        {
            $xmlTermLIVE.Name = $xmlTermSRC.Name
            $sMessageOut += "Name;"
        }

        #Term Description checks     
        if ($xmlTermLIVE.Descriptions.Description.Value.CompareTo($xmlTermSRC.Descriptions.Description.Value) -ne 0)
        {
            $xmlTermLIVE.Descriptions.Description.Value = $xmlTermSRC.Descriptions.Description.Value
            $xmlTermLIVE.Descriptions.Description.Language = $xmlTermSRC.Descriptions.Description.Language
            $sMessageOut += "Description;"
        }

        #Term LocalCustomProperties checks
        #Add any custom properties from the SRC file to the LIVE content
        # We only want to check the property UIN from SRC. There may be further properties added in SPO that we should not change.
        $xmlLocalCustomPropsSRC = $xmlTermSRC.SelectSingleNode("LocalCustomProperties")
        $xmlLocalCustomPropsLIVE = $xmlTermLIVE.SelectSingleNode("LocalCustomProperties")

        if ($xmlLocalCustomPropsLIVE.ChildNodes.Count -gt 0)
        {
            #LIVE has LocalCustomProperties, we need to check and amend
            if ($xmlTermLIVE.LocalCustomProperties.InnerXml.CompareTo($xmlTermSRC.LocalCustomProperties.InnerXml) -ne 0)
            {
                #We have issues with "See Also". In XSLT, we cannot combine strings on the fly, so the SRC input may have a
                # number of lines for "See Also" such as "See Also"/"See Also1"/"See Also2"/.../"See Also<n>"
                # We need to combine all of the "See Also" entries into a single value, ";" delimited. You cannot have multiple
                # properties with the same name. Se could just have multiple properties for "See Also" but it is neater to have them
                # all in a single property
                $sSeeAlsoValues = Get-SeeAlso($xmlLocalCustomPropsSRC.LocalCustomProperty)

                #Check to see if the KEY exists in LIVE
                foreach ($xmlLocalCustomPropSRC in $xmlLocalCustomPropsSRC.LocalCustomProperty)
                {
                    [boolean]$bAddElement = $true

                    #We want to skip any key that starts with "See Also" but does not end in "See Also"
                    # i.e. Skip "See Also1", "See Also2", etc.
                    if (($xmlLocalCustomPropSRC.Key.ToUpper().StartsWith("SEE ALSO")) -and (-not $xmlLocalCustomPropSRC.Key.ToUpper().EndsWith("SEE ALSO")))
                    {
                        $bAddElement = $false
                    }
                    elseif ($xmlLocalCustomPropSRC.Key.ToUpper() -eq "SEE ALSO")
                    {
                        $xmlSeeAlsoNode = $xmlLocalCustomPropsLIVE | select -ExpandProperty ChildNodes |where {$_.Key -eq "See Also"}

                        #See if the LIVE term already has the property. If not, we need to add it
                        if ($xmlSeeAlsoNode -ne $null)
                        {
                            #Property already exists in LIVe, so update
                            $xmlSeeAlsoNode.Value =  $sSeeAlsoValues
                            $bAddElement = $false
                        }
                    }
                    else
                    {
                        foreach ($xmlLocalCustomPropLIVE in $xmlLocalCustomPropsLIVE.ChildNodes)
                        {
                            if ($xmlLocalCustomPropLIVE.Attributes["Value"].Value -eq $xmlLocalCustomPropSRC.Attributes["Value"].Value)
                            {
                                $bAddElement = $false
                                break
                            }
                        }
                    }

                    if ($bAddElement)
                    {
                        $xmlElementLIVE = $xmlTermLIVE.OwnerDocument.CreateElement("LocalCustomProperty")

                        foreach ($xmlAttributeSRC in $xmlLocalCustomPropSRC.Attributes)
                        {
                            if ($xmlAttributeSRC.Name.ToUpper() -eq "SEE ALSO")
                            {
                                $xmlElementLIVE.setAttribute($xmlAttributeSRC.Name, $sSeeAlsoValues) > $null
                            }
                            else
                            {
                                $xmlElementLIVE.setAttribute($xmlAttributeSRC.Name, $xmlAttributeSRC.Value) > $null
                            }
                        }

                        $xmlLocalCustomPropsLIVE.AppendChild($xmlElementLIVE) > $null
                    }
                }
                $sMessageOut += "LocalCustomProperty;"
            }
        }
        else
        {
            #The LIVE term currently has no LocalCustomProperties
            $xmlElementLIVE = $xmlTermLIVE.OwnerDocument.CreateElement("LocalCustomProperty")

            foreach ($xmlAttributeSRC in $xmlLocalCustomPropsSRC.LocalCustomProperty.Attributes)
            {
                $xmlElementLIVE.setAttribute($xmlAttributeSRC.Name, $xmlAttributeSRC.Value) > $null
            }

            $xmlLocalCustomPropsLIVE.AppendChild($xmlElementLIVE) > $null
            $sMessageOut += "LocalCustomProperty;"
        }

        #Term Labels checks
        # First, see if we need to delve deeper
        if ($xmlTermLIVE.Labels.InnerXml.CompareTo($xmlTermSRC.Labels.InnerXml) -ne 0)
        {
            #Need to compare LIVE and SRC labels. First see if the property is already in LIVE, if not, just ADD the SRC one.
            $xmlLabelsSRC = $xmlTermSRC.SelectSingleNode("Labels")
            $xmlLabelsLIVE = $xmlTermLIVE.SelectSingleNode("Labels")

            #Check to see if the KEY exists in LIVE
            foreach ($xmlLabelSRC in $xmlLabelsSRC.Label)
            {
                [boolean]$bAddElement = $true

                foreach ($xmlLabelLIVE in $xmlLabelsLIVE.Label)
                {
                    if ($xmlLabelLIVE.Value.CompareTo($xmlLabelSRC.Value) -eq 0)
                    {
                        $bAddElement = $false
                        break
                    }
                }

                if ($bAddElement)
                {
                    #If there is a name change, then the default label needs to change, otherwise just add the new label
                    if ($xmlLabelSRC.IsDefaultForLanguage -eq $true)
                    {
                        #The first label is the default one
                        # We need to treat terms with a single label different from multiple.
                        # Multiple labels are returned as an array, singles are just a string
                        if ($xmlLabelsLIVE.ChildNodes.Count -eq 1)
                        {
                            $xmlLabelsLIVE.Label.Value = $xmlLabelSRC.Value
                        }
                        else
                        {
                            $xmlLabelsLIVE.Label[0].Value = $xmlLabelSRC.Value
                        }
                    }
                    else
                    {
                        $xmlElementLIVE = $xmlTermLIVE.OwnerDocument.CreateElement("Label")

                        foreach ($xmlAttributeSRC in $xmlLabelSRC.Attributes)
                        {
                            $xmlElementLIVE.setAttribute($xmlAttributeSRC.Name, $xmlAttributeSRC.Value) > $null
                        }

                        $xmlLabelsLIVE.AppendChild($xmlElementLIVE) > $null
                    }

                    #We may come through here a number of times but we only want to update the output string once for logging.
                    if ($sMessageOut.IndexOf("Labels;") -eq -1)
                    {
                        $sMessageOut += "Labels;"
                    }
                }
            }
        }

        #Final check. We need to check the parentnode for this term as it may have moved within the MMS structure
        if ($xmlTermLIVE.ParentNode.ParentNode.Name -ne $xmlTermSRC.ParentNode.ParentNode.Name)
        {
            $xPath = "//Term[translate(@Name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz') = """ + $xmlTermSRC.ParentNode.ParentNode.Name.ToLower() + """]"
            $xmlTermParentLIVE = $xmlTermsLIVE.SelectSingleNode($xPath)
            $xmlTermParentLIVE.ChildNodes[4].AppendChild($xmlTermLIVE) > $null
                   
            $sMessageOut += "Parent;"
        }

        Write-Host "Checks completed: " -ForegroundColor Green -NoNewline

        #Update the TERM Processed property so we know that is has been looked at
        # At the end, we can look for terms that haven't been processed to understand what's missing 
        # between RDM and the LIVE environment
        if ($sMessageOut -eq "")
        {
            $xmlTermLIVE.Processed = "No Changes"
            Write-Host "No modifications" -ForegroundColor White
        }
        else
        {
            $xmlTermLIVE.Processed = "Merged"
            Write-Host $sMessageOut -ForegroundColor Yellow
        }
    }
    catch
    {
        Write-Host "Hit a snag-$sMessageOut" -ForegroundColor Red
    }
}


################################################################
#
# Term-Path
# This function takes in an XML Term and returns a string of
# its heirarchy
################################################################
function Term-Path($xmlTerm) {

    $sTermPath = ""

    do
    {
        if ((-not ($xmlTerm.Name -eq "Terms")) -and  (-not ($xmlTerm.Name -eq "TermSets")))
        {
            $sTermPath = "\" + $xmlTerm.Name + $sTermPath
        }

        $xmlTerm = $xmlTerm.ParentNode
            
    } until ($xmlTerm.Name -eq "TermSets")

    return $sTermPath
}



################################################################
# This function works through the TERM nodes
#
################################################################
function FixUP-LIVETermsXML($xmlTermSRC, $xmlTermsLIVE, $global:glTermsModified, $global:glTermsAdded, $global:tsBusinessOwner, $global:glTermsCount){

    #Use this to keep track of changes made
    $sMessageOut = ""

    if ($xmlTermSRC.Name.ToLower().Contains("continuous improvement"))
    {
        Write-Host "here"
    }

    Write-Host "($global:glTermsCount of $global:glTermsApprox [est]) Processing Term " -ForegroundColor White -NoNewline
    Write-Host $xmlTermSRC.Name -ForegroundColor Cyan -NoNewline
    Write-Host " ... " -ForegroundColor White -NoNewline

    #We need to figure out the XPath of the current node to search the LIVE XML against for matching
    # We can use a do/while to navigate up to build the XPath but put a capture
    $xPath = "//Term[translate(@Name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz') = """ + $xmlTermSRC.Name.ToLower().Trim() + """]"

    $xmlTermsMultiLIVE = $xmlTermsLIVE.SelectNodes($xPath)

    #In case the "Name" property is different, just reselect using the CVRObjectID
    if ($xmlTermsMultiLIVE.Count -eq 0)
    {
        #We should recheck using the CVRObjectID
        $xPath = "//Term[translate(@CVRObjectID,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz') = " + $xmlTermSRC.CVRObjectID + "]"
        $xmlTermsMultiLIVE = $xmlTermsLIVE.SelectNodes($xPath)
    }

    if ($xmlTermsMultiLIVE.Count -ne 0)
    {
        #This should only give one return but just in-case, check
        if ($xmlTermsMultiLIVE.Count -eq 1)
        {
            $xmlTermLIVE = $xmlTermsLIVE.SelectSingleNode($xPath)

            Write-Host "Comparing" -ForegroundColor DarkYellow -NoNewline
            Write-Host " ... " -ForegroundColor White -NoNewline

            Compare-Terms $xmlTermsLIVE $xmlTermLIVE $xmlTermSRC $sMessageOut

            ++$global:glTermsModified
        }
        else
        {
            #Found multiple matches.
            Write-Host "Duplicate" -ForegroundColor Red -NoNewline
            Write-Host " ... " -ForegroundColor White -NoNewline

            #We have duplicate terms. We need to see which is correct and offer the removal of duplicate terms
            # Find the correct term from LIVE then provide that for comparison
            # Offer the duplicate term/s for Fixup

            #Get the SRC term full path
            $sTermPathSRC = Term-Path($xmlTermSRC)

            foreach ($xmlTermLIVE in $xmlTermsMultiLIVE)
            {
                $sTermPathLIVE = Term-Path($xmlTermLIVE)

                if ($sTermPathSRC.ToLower().Trim() -eq $sTermPathLIVE.ToLower().Trim())
                {
                    #We have found the correct term. The SRC is the correct entry
                    # Check and fixup that term
                    Compare-Terms $xmlTermsLIVE $xmlTermLIVE $xmlTermSRC $sMessageOut
                    ++$global:glTermsModified
                }
                else
                {
                    $xmlTermLIVE.Processed = "Merged-Duplicate"
                }
            }
        }
    }
    else
    {
        #New term not deployed in TermStore
        Write-Host "Adding" -ForegroundColor Yellow -NoNewline
        Write-Host " ... " -ForegroundColor White -NoNewline

        #Find the parent of the current SRC term in the LIVE XML. It should always exist.
        $bRootNode = $false

        #If a root node, then we only want ParentNode, otherwise, ParentNode.ParentNode
        if ($xmlTermSRC.ParentNode.ParentNode.ParentNode.ParentNode.ParentNode.ParentNode.ParentNode.ParentNode.LocalName -eq "#document")
        {
            $xmlTermParentLIVE = $xmlTermsLIVE
            $bRootNode = $true
        }
        else
        {
            $xPath = "//Term[translate(@Name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz') = """ + $xmlTermSRC.ParentNode.ParentNode.Name.ToLower() + """]"
            $xmlTermParentLIVE = $xmlTermsLIVE.SelectSingleNode($xPath)
        }

        #Check to ensure we can place it somewhere
        if ($xmlTermParentLIVE -ne $null)
        {
            $xmlTermLIVE = $xmlTermParentLIVE.OwnerDocument.CreateElement("Term")
            $xmlTermLIVE.setAttribute("Name", $xmlTermSRC.Name)
            $xmlTermLIVE.setAttribute("Id", "")
            $xmlTermLIVE.setAttribute("CustomSortOrder", "")
            $xmlTermLIVE.setAttribute("IsAvailableForTagging", "True")
            $xmlTermLIVE.setAttribute("Processed", "False")
            $xmlTermLIVE.setAttribute("CVRObjectID", $xmlTermSRC.CVRObjectID)
            $xmlTermLIVE.setAttribute("IsSourceTerm", "True")
            $xmlTermLIVE.setAttribute("IsRoot", "False")
            $xmlTermLIVE.setAttribute("IsReused", "False")
            $xmlTermLIVE.setAttribute("IsKeyword", "False")
            $xmlTermLIVE.setAttribute("IsDeprecated", "False")

            if (-not $bRootNode) {$xmlTermParentLIVE = $xmlTermParentLIVE.SelectSingleNode("Terms")}
            $xmlTermLIVE = $xmlTermParentLIVE.AppendChild($xmlTermLIVE)

            #Need to add DESCRIPTION
            $xmlDescriptionsLIVE = $xmlTermLIVE.OwnerDocument.CreateElement("Descriptions")
            $xmlDescriptionsLIVE = $xmlTermLIVE.AppendChild($xmlDescriptionsLIVE)

            $xmlDescriptionLIVE = $xmlDescriptionsLIVE.OwnerDocument.CreateElement("Description")
            $xmlDescriptionLIVE.SetAttribute("Value", $xmlTermSRC.Descriptions.Description.Value) > $null
            $xmlDescriptionLIVE.SetAttribute("Language", "1033") > $null
            $xmlDescriptionsLIVE.AppendChild($xmlDescriptionLIVE) > $null

            #Need to add CUSTOMPROPERTIES node
            $xmlCustomPropsLIVE = $xmlTermLIVE.OwnerDocument.CreateElement("CustomProperties")
            $xmlCustomPropsLIVE = $xmlTermLIVE.AppendChild($xmlCustomPropsLIVE)

            #Need to add a LOCALCUSTOMPROPERTIES node
            $xmlLocalCustomPropsSRC = $xmlTermSRC.SelectSingleNode("LocalCustomProperties")
            $xmlLocalCustomPropsLIVE = $xmlTermLIVE.OwnerDocument.CreateElement("LocalCustomProperties")
            $xmlLocalCustomPropsLIVE = $xmlTermLIVE.AppendChild($xmlLocalCustomPropsLIVE)

            foreach ($xmlLocalCustomPropSRC in $xmlLocalCustomPropsSRC.LocalCustomProperty)
            {
                if (-not (($xmlLocalCustomPropSRC.Key.ToUpper().StartsWith("SEE ALSO")) -and ($xmlLocalCustomPropSRC.Key.Length -gt 8)))
                {
                    $xmlLocalCustomPropLIVE = $xmlLocalCustomPropsLIVE.OwnerDocument.CreateElement("LocalCustomProperty")
               
                    #Copy the key...followed by the value for the key
                    $xmlLocalCustomPropLIVE.SetAttribute("Key", $xmlLocalCustomPropSRC.Key) > $null

                    if ($xmlLocalCustomPropSRC.Key.ToUpper() -eq "SEE ALSO")
                    {
                        #We have to manage the "See Also" LocalCustomProperty differently
                        $sSeeAlsoValues = Get-SeeAlso($xmlLocalCustomPropsSRC)
            
                        #Not all terms have a "See Also" node but if we have found one, chop off the last ";" character
                        if ($sSeeAlsoValues.Length -gt 0)
                        {
                            $xmlLocalCustomPropLIVE.SetAttribute("Value", $sSeeAlsoValues) > $null
                        }
                    }
                    else
                    {
                        $xmlLocalCustomPropLIVE.SetAttribute("Value", $xmlLocalCustomPropSRC.Value) > $null
                    }
                    $xmlLocalCustomPropsLIVE.AppendChild($xmlLocalCustomPropLIVE) > $null
                }
            }

            #Need to add a LABELS node
            $xmlLabelsSRC = $xmlTermSRC.SelectSingleNode("Labels")
            $xmlLabelsLIVE = $xmlTermLIVE.OwnerDocument.CreateElement("Labels")
            $xmlLabelsLIVE = $xmlTermLIVE.AppendChild($xmlLabelsLIVE)

            foreach ($xmlLabelSRC in $xmlLabelsSRC.ChildNodes)
            {
                $xmlLabelLIVE = $xmlLabelsLIVE.OwnerDocument.CreateElement("Label")

                foreach ($xmlAttributeSRC in $xmlLabelSRC.Attributes)
                {
                    $xmlLabelLIVE.setAttribute($xmlAttributeSRC.Name, $xmlAttributeSRC.Value) > $null
                }
                $xmlLabelsLIVE.AppendChild($xmlLabelLIVE) > $null
            }

            #Need to add a TERMS node
            $xmlElementLIVE = $xmlTermLIVE.OwnerDocument.CreateElement("Terms")
            $xmlElementsLIVE = $xmlTermLIVE.AppendChild($xmlElementLIVE)

            #Update the TERM Processed property so we know that is has been looked at
            # At the end, we can look for terms that haven't been processed to understand what's missing 
            # between SRC and the LIVE environment
            $xmlTermLIVE.Processed = "Merged"

            Write-Host "Completed:" -ForegroundColor Green

            ++$global:glTermsAdded
        }
        else
        {
            #Can't find a parent for the new term. Skip
            Write-Host "Invalid" -ForegroundColor Red
        }
    }

    ++$global:glTermsCount

    #Process the childnodes here
    foreach ($xmlChildTerm in $xmlTermSRC.Terms.ChildNodes)
    {
        FixUP-LIVETermsXML $xmlChildTerm $xmlTermsLIVE $global:glTermsModified $global:glTermsAdded $global:tsBusinessOwner $global:glTermsCount
    }
}


#####################################################################################################
#
# This is where the code will run.
#
#####################################################################################################

#Inform the user how long the operation took.  This is the start date/time.
$dtStartTime = Get-Date

Write-Host "##################################################################" -ForegroundColor Cyan
Write-Host "Starting the MMS File Checker at $dtStartTime" -ForegroundColor Cyan
Write-Host "** Please note that the XML files to be merged must both be     **" -ForegroundColor Cyan
Write-Host "** located in the same directory                                **" -ForegroundColor Cyan
Write-Host "##################################################################" -ForegroundColor Cyan

#Ask for the directory path then make sure it is valid
if ($FilePath -eq "") {$FilePath = Read-Host "Which directory contains the XML files to be checked"}
if (-not (Test-Path $FilePath))
{
    do
    {
        $FilePath = Read-Host "Cannot find the directory [$FilePath]. Please RE-ENTER"
    } while (-not (Test-Path $FilePath))
}

#Ask for the SRC Transformed file name and make sure it is valid
if (($SRCExport -eq "") -or ($SRCExport -eq $null)) {$SRCExport = Read-Host "What is the name of the TRANSFORMED XML file"}
$sFilePathOfXMLSRC = $FilePath + "\" + $SRCExport

if (-not (Test-Path $sFilePathOfXMLSRC))
{
    do
    {
        $SRCExport = Read-Host "Cannot find the SRC Transformed file [$SRCExport]. Please RE-ENTER"
        $sFilePathOfXMLSRC = $FilePath + "\" + $SRCExport
    } while (-not (Test-Path $sFilePathOfXMLSRC))
}

#Ask for the LIVE Export file name and make sure it is valid
if (($LIVEExport -eq "") -or ($LIVEExport -eq $null)) {$LIVEExport = Read-Host "What is the name of the TENANT EXPORT XML file"}
$sFilePathOfXMLLIVE = $FilePath + "\" + $LIVEExport

if (-not (Test-Path $sFilePathOfXMLLIVE))
{
    do
    {
        $LIVEExport = Read-Host "Cannot find the LIVE Export file [$LIVEExport]. Please RE-ENTER"
        $sFilePathOfXMLLIVE = $FilePath + "\" + $LIVEExport
    } while (-not (Test-Path $sFilePathOfXMLLIVE))
}

#Set a boolean as we need to merge the Business Onwer termset differently
if ($SRCExport.IndexOf("BusinessOwner-ATLAS") -ge 0)
{
    #Checking Business Owner
    $global:tsBusinessOwner = $true
}

#We have issues with some non-supported characters in SPO MMS. So, normalise the SRC file first.
# It is only in the SRC file that we can get unsupported characters as the Live file is exported from SPO
# Better to do it here otherwise we will have to manage each items we compare separately.
$sFilePathOfXMLSRC = $FilePath + "\" + $SRCExport
Write-Host "Please Wait... We need to remove unsupported characters from " -NoNewline -ForegroundColor Cyan 
Write-Host $sFilePathOfXMLSRC -ForegroundColor Green
Write-Host "This may take a few minutes."  -ForegroundColor Cyan

#The filename is also updated so we don't have to re-process this file
$sFilePathOfXMLSRC = Cleanse-File $sFilePathOfXMLSRC

#Get the main working files
$xmlDocumentSRC = [xml](Get-Content $sFilePathOfXMLSRC)
$xmlDocumentLIVE = [xml](Get-Content $sFilePathOfXMLLIVE)

$sDate = (Get-Date).ToString("yyyyMMdd")
$sFilePathOfXMLFixup = $FilePath + "\" + $sDate + "-" + $xmlDocumentLIVE.TermStores.TermStore.Groups.Group.Name.Replace(" ", "") + "-MMS-Fixup.XML"

Write-Host "This run is looking at the group [" -NoNewline -ForegroundColor Cyan
Write-Host $xmlDocumentLIVE.TermStores.TermStore.Groups.Group.Name -NoNewline -ForegroundColor Green
Write-Host "] and the termset [" -NoNewline -ForegroundColor Cyan
Write-Host $xmlDocumentLIVE.TermStores.TermStore.Groups.Group.TermSets.TermSet.Name -NoNewline -ForegroundColor Green
Write-Host "]" -ForegroundColor Cyan
Write-Host "##################################################################" -ForegroundColor Cyan

#Make sure we have the correct Transformed (SRC) file. We can tell from the Group ID which should be "TRANSFORMED"
if (($xmlDocumentSRC.TermStores.TermStore.Groups.Group.ID.ToUpper().CompareTo("TRANSFORMED") -eq 0) -and ($xmlDocumentLIVE.TermStores.TermStore.Groups.Group.ID.ToUpper().CompareTo("TRANSFORMED") -ne 0))
{
    #Check to make sure we have the correct files by looking at the Termset names
    if ($xmlDocumentLIVE.TermStores.TermStore.Groups.Group.Name.CompareTo($xmlDocumentSRC.TermStores.TermStore.Groups.Group.Name) -ne 0)
    {
        #Work out the elapsed time.
        $dtEndTime = Get-Date
        $dtDifference = New-TimeSpan -Start $dtStartTime -End $dtEndTime
        Write-Host ""
        Write-Host "The Source file " -NoNewline -ForegroundColor White
        Write-Host $SRCExport -NoNewline -ForegroundColor Green
        Write-Host ", TermSet = [" -NoNewline -ForegroundColor White
        Write-Host $xmlDocumentLIVE.TermStores.TermStore.Groups.Group.Name -NoNewline -ForegroundColor Green
        Write-Host "]" -ForegroundColor White
        Write-Host "does not match the termset defined in " -NoNewline -ForegroundColor White
        Write-Host $LIVEExport -NoNewline -ForegroundColor Green
        Write-Host ", TermSet = [" -NoNewline -ForegroundColor White
        Write-Host $xmlDocumentSRC.TermStores.TermStore.Groups.Group.Name -NoNewline -ForegroundColor Green
        Write-Host "]" -ForegroundColor White
        Write-Host ""
        Write-Host "These 2 files cannot be merged. Please retry with the correct files for matching termsets." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "#############################################################" -ForegroundColor Cyan
    }
    else
    {
        #Ready now to process the required changes. We pass all of the LIVE terms as it is only used for searching against
        $xmlTermsSRC = $xmlDocumentSRC.SelectSingleNode("//Terms")
        $xmlTermsLIVE = $xmlDocumentLIVE.SelectSingleNode("//Terms")

        #Get the approx term count (Approx as we don't know how many "new" terms will come from the source
        $global:glTermsApprox = $xmlDocumentLIVE.TermStores.TermStore.Groups.Group.TermSets.TermSet.TermCount

        foreach ($xmlTermSRC in $xmlTermsSRC.Term)
        {
            FixUP-LIVETermsXML $xmlTermSRC $xmlTermsLIVE $global:glTermsModified $global:glTermsAdded $global:tsBusinessOwner $global:glTermsCount
            ++$global:glTermsCount
        }

        #We need to update the TermSet count as new items may have been added to the original export
        $xmlDocumentLIVE.TermStores.TermStore.Groups.Group.TermSets.TermSet.TermCount = GetTerm-Count($xmlDocumentLIVE)

        #Final Checks
        # Read the LIVE term file and check to see if every term has been "Processed=True"
        Check-Processed $xmlDocumentLIVE $sFilePathOfXMLLIVE

        #The LIVE input XML will be updated with all of the SRC content, ready for re-import
        $xmlDocumentLIVE.Save($sFilePathOfXMLLIVE.Replace(".XML", "-SRC_PROCESSED.XML"))

        $sFilePathOfXMLLIVEProcessed = $sFilePathOfXMLLIVE.Replace(".XML", "-SRC_PROCESSED.xml")

        #Work out the elapsed time.
        $dtEndTime = Get-Date
        $dtDifference = New-TimeSpan -Start $dtStartTime -End $dtEndTime
        Write-Host ""
        Write-Host "Please use [$sFilePathOfXMLLIVEProcessed] and [ImportMMSGroup.PS1] to update the MMS group ["$xmlDocumentLIVE.TermStores.TermStore.Groups.Group.Name"]" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "#############################################################" -ForegroundColor Cyan
        Write-Host " $global:glTermsModified terms were MODIFIED" -ForegroundColor Cyan
        Write-Host " $global:glTermsAdded new terms were ADDED" -ForegroundColor Cyan
    }
}
else
{
    #We have the wrong SRC file
    
    #Work out the elapsed time.
    $dtEndTime = Get-Date
    $dtDifference = New-TimeSpan -Start $dtStartTime -End $dtEndTime
    Write-Host ""

    if ($xmlDocumentSRC.TermStores.TermStore.Groups.Group.ID.ToUpper().CompareTo("TRANSFORMED") -ne 0)
    {
        Write-Host ""
        Write-Host "The Source file " -NoNewline -ForegroundColor White
        Write-Host $SRCExport -NoNewline -ForegroundColor Green
        Write-Host ", is not a valid Transformed SRC file!" -ForegroundColor White
        Write-Host "The Group ID should equal [" -NoNewline -ForegroundColor White
        Write-Host "TRANSFORMED" -NoNewline -ForegroundColor Green
        Write-Host "]" -NoNewline -ForegroundColor White
        Write-Host " but it is set as [" -NoNewline -ForegroundColor White
        Write-Host $xmlDocumentSRC.TermStores.TermStore.Groups.Group.Name -NoNewline -ForegroundColor Red
        Write-Host "]" -ForegroundColor White
        Write-Host ""
        Write-Host "These 2 files cannot be merged. Please retry with the correct files for matching termsets." -ForegroundColor Yellow
        Write-Host "Probable issue is that you have entered the Export file as the Transformed version" -ForegroundColor Yellow
    }

    if ($xmlDocumentLIVE.TermStores.TermStore.Groups.Group.ID.ToUpper().CompareTo("TRANSFORMED") -eq 0)
    {
        Write-Host "The Export file " -NoNewline -ForegroundColor White
        Write-Host $LIVEExport -NoNewline -ForegroundColor Green
        Write-Host ", is not a valid Live export file!" -ForegroundColor White
        Write-Host "The Group ID should not equal [" -NoNewline -ForegroundColor White
        Write-Host "TRANSFORMED" -NoNewline -ForegroundColor Green
        Write-Host "]" -NoNewline -ForegroundColor White
        Write-Host " but it is set as [" -NoNewline -ForegroundColor White
        Write-Host $xmlDocumentSRC.TermStores.TermStore.Groups.Group.Name -NoNewline -ForegroundColor Red
        Write-Host "]" -ForegroundColor White
        Write-Host ""
        Write-Host "These 2 files cannot be merged. Please retry with the correct files for matching termsets." -ForegroundColor Yellow
        Write-Host "Probable issue is that you have entered the Transformed file as the Export version" -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "#############################################################" -ForegroundColor Cyan
}

#Finish off
Write-Host " This took $dtDifference to complete" -ForegroundColor Cyan
Write-host " Completed..." -ForegroundColor Cyan
Write-Host "#############################################################" -ForegroundColor Cyan
#End Of script