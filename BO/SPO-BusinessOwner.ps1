#############################################################################################################
# 
# Name:
#       SPO-BusinessOwner.PS1
#
# Version:
#       V1.0
#
# Description:
#       This PowerShell script can be used to convert the Business Owner structure as delivered by ATLAS
#       called "EUN.xls" into a form that the current tooling "MergeXMLs.PS1", can understand.
#
#       This replicates the work of the SPO-SubjectCategory.XSLT and SPO-SubjectKeywords.XSLT
#       transformation files. Those XSLT file takes a single XML file provided by DBS that determines the
#       UK Defence Taxonomy and Thesaurus structures. However, the Business Owner termset is derived from
#       a different source owned by ATLAS and maintained by the EUN database. Typically, this database is
#       only intended to provision role names but as these are no longer relevant, there are questions as
#       to whether that should remain as the master source. We should now be treating SPO as the master
#       source of the EUN structure.
#
#       Before this script can be run, that spreadsheet (EUN.XLS) has to be converted to a CSV file. 
#       In order to do this, perform the following steps:
#         1) Open "EUN.xls"" as provided by ATLAS
#         2) Select "File"/"Save As" and select a file extension of "CSV (Comma delimited)(*.csv)"
#         3) Save the file.
#       This routine will then create a date/time stamped <yyyymmmdd>-BusinessOwner-ATLAS.XML file that
#       can be fed into "MergeXMLs.PS1" with a tenant exported "Business Owner" XML produced by 
#       "ExportMMSGroup.PS1"
#
#       NOTE: This script can only operate with CSV files, not Excel original files.
#
# Author: 
#       Nigel Bridport, Microsoft Consulting Services
#
# Date: 
#       06/10/2016
#
# Changes:
#       Version  Date      Who             Description
#       V1.0     06/10/16  Nigel Bridport  Initial file creation
#
# Inputs:
#       $FilePath - Full file path to, and including, the filename
#
# Example:
#       PS> SPO-BusinessOwner.PS1 -FilePath "C:\Users\nigelbri\PS Scripts\Taxonomy Management\EUN.csv"
#
#############################################################################################################
#
Param(
    [Parameter(Mandatory=$false)]
    [string]$FilePath
)

#Can change the line below to control the output filename. By default, this will be "<yyymmdd>-BusinessOwner-ATLAS.XML"
$gOutputFilePostfix = "-BusinessOwner-ATLAS.XML"

#####################################################################################################
#
# Get-XMLTermStoreTemplateToFile
# This function creates a temporary template file that is used to populate information from the  
# supplied ATLAS EUN csv file. 
#
#####################################################################################################
function Get-XMLTermStoreTemplateToFile($sPath){
    #Set up an xml template used for creating an exported xml
    $xmlTemplate = '<TermStores>
                      <TermStore>
                        <Groups>
                          <Group Name="Business Owner" IsSiteCollectionGroup="" IsSystemGroup="" Description="" Id="TRANSFORMED">
                            <TermSets>
                              <TermSet Name="Business Owner" Description="" Contact="" IsAvailableForTagging="" TermCount="">
                                <Terms>
                                  <Term Id="" Name="" IsDeprecated="" IsAvailableForTagging="" IsKeyword="" IsReused="" IsRoot="" IsSourceTerm="" CustomSortOrder="" CVRObjectID="" Processed="False">
                                    <Descriptions>
                                      <Description Language="1033" Value="" />
                                    </Descriptions>
                                    <CustomProperties>
                                      <CustomProperty Key="" Value="" />
                                    </CustomProperties>
                                    <LocalCustomProperties>
                                      <LocalCustomProperty Key="" Value="" />
                                    </LocalCustomProperties>
                                    <Labels>
                                      <Label Value="" Language="1033" IsDefaultForLanguage="" />
                                    </Labels>
                                    <Terms>
                                      <Term Id="" Name="" IsDeprecated="" IsAvailableForTagging="" IsKeyword="" IsReused="" IsRoot="" IsSourceTerm="" CustomSortOrder="" CVRObjectID="" Processed="False">
                                        <Descriptions>
                                          <Description Language="1033" Value="" />
                                        </Descriptions>
                                        <CustomProperties>
                                          <CustomProperty Key="" Value="" />
                                        </CustomProperties>
                                        <LocalCustomProperties>
                                          <LocalCustomProperty Key="" Value="" />
                                        </LocalCustomProperties>
                                        <Labels>
                                          <Label Value="" Language="1033" IsDefaultForLanguage="" />
                                        </Labels>
                                      </Term>
                                    </Terms>
                                  </Term>
                                </Terms>
                              </TermSet>
                            </TermSets>
                          </Group>
                        </Groups>
                      </TermStore>
                    </TermStores>'    

    try
    {
        #Save Template to disk
        $xmlTemplate | Out-File($sPath + "\Template.xml")
 
        #Load file and return
        $xml = New-Object XML
        $xml.Load($sPath + "\Template.xml")
        return $xml
    }
    catch
    {
        Write-Host "Error creating Template file. " $_.Exception.Message -ForegroundColor Red
        exit 1
    }
}


#####################################################################################################
#
# Get-XMLFileObjectTemplates 
# This function creates the templates nodes are loaded for replacement as we enumerate the required
# taxonomy structure.
#
#####################################################################################################
function Get-XMLFileObjectTemplates($xml){
    $global:xmlTermSetT = $xml.selectSingleNode('//TermSet[@Id=""]')  
    $global:xmlTermT = $xml.selectSingleNode('//Term[@Id=""]')
    $global:xmlTermLabelT = $xml.selectSingleNode('//Label[@Value=""]')
    $global:xmlTermDescriptionT = $xml.selectSingleNode('//Description[@Value=""]')
    $global:xmlTermCustomPropertiesT = $xml.selectSingleNode('//CustomProperty[@Key=""]')
    $global:xmlTermLocalCustomPropertiesT = $xml.selectSingleNode('//LocalCustomProperty[@Key=""]')
}


#####################################################################################################
#
# Clean-Template 
# This function ensures that the templates are clean and empty before starting to enumerate the 
# taxonomy with data
#
#####################################################################################################
function Clean-Template($xml) {
    #Do not cleanup empty description nodes (this is the default state)

    #Empty Term.Labels.Label
    $xml.SelectNodes('//Label[@Value=""]') | ForEach-Object {
        $parent = $_.get_ParentNode()
        $parent.RemoveChild($_)  | Out-Null      
    } 
    
    #Empty Term
    $xml.SelectNodes('//Term[@Id=""]') | ForEach-Object {
        $parent = $_.get_ParentNode()
        $parent.RemoveChild($_)  | Out-Null      
    } 
    
    #Empty TermSet
    $xml.SelectNodes('//TermSet[@Id=""]') | ForEach-Object {
        $parent = $_.get_ParentNode()
        $parent.RemoveChild($_)  | Out-Null      
    }
    
    #Empty Custom Properties
    $xml.SelectNodes('//CustomProperty[@Key=""]') | ForEach-Object {
        $parent = $_.get_ParentNode()
        $parent.RemoveChild($_) | Out-Null
    }

    #Empty Local Custom proeprties
    $xml.SelectNodes('//LocalCustomProperty[@Key=""]') | ForEach-Object {
        $parent = $_.get_ParentNode()
        $parent.RemoveChild($_) | Out-Null
    }

    $xml.SelectNodes('//Descriptions')| ForEach-Object {
        $childNodes = $_.ChildNodes.Count
        
        if ($childNodes -gt 1)
        {
            $_.RemoveChild($_.ChildNodes[0]) | Out-Null
        }
    }

    while ($xml.SelectNodes('//Term[@Id=""]').Count -gt 0)
    {
        #Cleanup the XML, remove empty Term Nodes
        $xml.SelectNodes('//Term[@Id=""]').RemoveAll() | Out-Null
    }
}

   
#####################################################################################################
#
# AddTerm 
# This function takes input data that includes the parent term and current term, and then adds that
# information into the correct location in the XML domain object.
#
#####################################################################################################
function AddTerm($xmlOutputFile, $csvTLB, $xmlParent, $oEUN, $global:lTermCount, $global:lTermCountTotal){

    #Get the term template object and then populate
    $xmlTerm = $global:xmlTermT.Clone()
    $xmlTerm.Name = $oEUN.'Electronic Unit Name'
    $xmlTerm.Descriptions.Description.Value = $oEUN.'Long Unit Name'
    if (($oEUN.'With Effect From' -ne $null) -or ($oEUN.'With Effect From' -ne ""))
    {
        $xmlTerm.Descriptions.Description.Value += ". Unit name effective from " + $oEUN.'With Effect From'
    }

    #Add the UIN, Financial Information, as a LocalCustomProperty called "UIN".
    $xmlTerm.LocalCustomProperties.LocalCustomProperty.Key = "UIN"
    $xmlTerm.LocalCustomProperties.LocalCustomProperty.Value = $oEUN.UIN

    #Add the labels.
    #  First label is by default always the same as the name
    #  Second label is the "Long Unit Name" - Need to ensure that the "Long Unit Name" is not the same as the default label. This can happen.
    #  Deprecated labels are stored in the "Legacy EUN(s)" property. This property is ";" delimited
    $sEUNDefaultLabels = $oEUN.'Electronic Unit Name'.Trim() + ";"

    if ($oEUN.'Electronic Unit Name'.Trim().ToLower() -ne $oEUN.'Long Unit Name'.Trim().ToLower())
    {
        $sEUNDefaultLabels = $sEUNDefaultLabels + $oEUN.'Long Unit Name'.Trim() + ";"
    }

    $aLabels = ($sEUNDefaultLabels + $oEUN.'Legacy EUN(s)').TrimEnd(";").Split(";")

    #This counter is used as occasionally the "Labels" node is seen as a string and not an XMLLinkedNode
    $iCount = 1

    #Enumerate through the labels and add to the labels node
    foreach ($sLabel in $aLabels)
    {
        #Clone Term Label node
        $xmlTermLabel = $global:xmlTermLabelT.Clone()
        $xmlTermLabel.Value = $sLabel
        
        if ($iCount -eq 1) 
        {
            $xmlTermLabel.IsDefaultForLanguage = "True"

            #HACK: Problem with the object for labels. As a workaround, directly point at the index value but should revisit later.
            $xmlTerm.ChildNodes[3].AppendChild($xmlTermLabel) | Out-Null
        }
        else
        {
            $xmlTermLabel.IsDefaultForLanguage = "False"
            $xmlTerm.Labels.AppendChild($xmlTermLabel) | Out-Null
        }
        $iCount++
    }
 
    #This is the TLB root node added. Now for the children
    $xmlParent = $xmlParent.AppendChild($xmlTerm)

    #Look in the current TLB, for any children terms to the just added term. If we find any,
    # we will need to enumerate them now before moving to the next term at this items peer level.
    $oEUNs = $csvTLB.Group | where {$_.'Parent EUN' -eq $oEUN.'Electronic Unit Name'}

    #Enumerate each TLB and write out to the XML file
    foreach ($oEUN in $oEUNs)
    {
        $global:lTermCount++
        Write-Host "($global:lTermCount of $global:lTermCountTotal) " -ForegroundColor White -NoNewline
        Write-Host "Creating the [" -ForegroundColor Cyan -NoNewline
        Write-Host $oEUN.'Electronic Unit Name' -ForegroundColor Yellow -NoNewline
        Write-Host "] unit underneath [" -ForegroundColor Cyan -NoNewline
        Write-Host $oEUN.'Parent EUN' -ForegroundColor Yellow -NoNewline
        Write-Host "]" -ForegroundColor Cyan
        AddTerm $xmlOutputFile $csvTLB $xmlParent.ChildNodes[4] $oEUN $global:lTermCount $global:lTermCountTotal
    }
}


#####################################################################################################
#
# SPO-BusinessOwner.PS1
# This is where the main routine will run from.
#
#####################################################################################################
#

#Inform the user how long the operation took.  This is the start date/time.
$dtStartTime = Get-Date

Write-Host "##################################################################" -ForegroundColor Cyan
Write-Host "Starting the Business Owner conversion at $dtStartTime" -ForegroundColor Cyan
Write-Host "##################################################################" -ForegroundColor Cyan

#Check to make sure the entered filepath exists
if (($FilePath -eq "") -or ($FilePath -eq $null)) {$FilePath = Read-Host "What is the filepath to the Business Owner csv"}
if (-not (Test-Path $FilePath))
{
    do
    {
        $FilePath = Read-Host "Cannot find the CSV file [$FilePath]. Please RE-ENTER"
    } while (-not (Test-Path $FilePath))
}

#The date information is used to create the output files.
$sDate = (Get-Date).ToString("yyyyMMdd")
$sDirPath = $FilePath.Substring(0, $FilePath.LastIndexOf("\"))
$sXMLFixup = $sDate + $gOutputFilePostfix

#Get the CSV file data
$csvTLBs = Import-Csv $FilePath | Group {$_.TLB}

#Create the XML template and ready the XML node objects
$xmlOutputFile = Get-XMLTermStoreTemplateToFile $sDirPath
Get-XMLFileObjectTemplates $xmlOutputFile
Clean-Template $xmlOutputFile

#Some counters
$iTLBTotalCount = $csvTLBs.Count
$iTLBCurrentCount = 1
$iTotalTerms = 0

#Enumerate each TLB and write out to the XML output file
foreach ($csvTLB in $csvTLBs)
{
    $iTotalTerms += $csvTLB.Count
    Write-Host "($iTLBCurrentCount of $iTLBTotalCount): " -ForegroundColor White -NoNewline
    Write-Host "Creating the TLB organisation structure for [" -ForegroundColor Cyan -NoNewline
    Write-Host $csvTLB.Name -ForegroundColor Yellow -NoNewline
    Write-Host "]" -ForegroundColor Cyan

    #Keep a note of the total terms within the TLB
    $global:lTermCountTotal = $csvTLB.Count

    #Get the TLB EUN entry
    $oEUN = $csvTLB.Group | where {$_.'Electronic Unit Name' -eq $csvTLB.Name}

    #The Excel file is not consistent. There are occasions where there is no EUN entry for a TLB, i.e. NC.
    # To get around that issue, we may need to create a temporary object as we know what TLBS are in the export
    if ($oEUN -eq $null) 
    {
        $oEUN = New-Object psobject -Property @{"Electronic Unit Name"=""; "Long Unit Name"=""; "Parent EUN"=""; "Bowmanised"=""; "Legacy EUN(s)"=""; "UIN"=""; "With Effect From"=""; "TLB"=""}
        $oEUN.'Electronic Unit Name' = $csvTLB.Name

        #If we create an EUN for a TLB, as there is an issue with the EUN db in that not all TLBs have to have an EUN, we need to add 1 to the total EUN count for the display counter to look correct
        $global:lTermCountTotal += 1
    }

    $global:lTermCount = 1

    Write-Host $global:lTermCount" of "$global:lTermCountTotal -ForegroundColor Yellow -NoNewline
    Write-Host " : Creating the [" -ForegroundColor White -NoNewline
    Write-Host $oEUN.'Electronic Unit Name' -ForegroundColor Yellow -NoNewline
    Write-Host "] unit. TLB node." -ForegroundColor White

    AddTerm $xmlOutputFile $csvTLB $xmlOutputFile.TermStores.TermStore.Groups.Group.TermSets.TermSet.ChildNodes[0] $oEUN 1 $global:lTermCountTotal

    $iTLBCurrentCount++
}

#Put the total term count into the XML output file. (TermSet node)
$xmlOutputFile.TermStores.TermStore.Groups.Group.TermSets.TermSet.TermCount = $iTotalTerms.ToString()

#Save file.
try
{
    $xmlOutputFile.Save($sDirPath + "\NewTaxonomy.xml")
   
    #Clean up empty <Term> unable to work out in Clean-Template.
    Get-Content ($sDirPath + "\NewTaxonomy.xml") | Foreach-Object { $_ -replace "<Term><\/Term>", "" } | Set-Content ($sDirPath + "\" + $sXMLFixup)

    #Remove temp file
    Remove-Item($sDirPath + "\Template.xml");
    Remove-Item($sDirPath + "\NewTaxonomy.xml");
}
catch
{
    Write-Host "Error saving XML File to disk " $_.Exception.Message -ForegroundColor Red
    exit 1
}
finally
{
    #Work out the elapsed time.
    $dtEndTime = Get-Date
    $dtDifference = New-TimeSpan -Start $dtStartTime -End $dtEndTime

    Write-Host "##################################################################" -ForegroundColor Cyan
    Write-Host "Completed..." -ForegroundColor Cyan
    Write-Host "The CSV file [$FilePath]" -ForegroundColor Cyan 
    Write-Host "has been transformed into [$sDirPath\$sXMLFixup]" -ForegroundColor Cyan 
    Write-Host "which took $dtDifference" -ForegroundColor Cyan
    Write-Host "##################################################################" -ForegroundColor Cyan
}