###############################################################################################################
# 
#  NAME:  ExportMMSGroup.PS1
#
#  VERSION: V1.2
# 
#  DATE: 12th October 2016
#
#  DESCRIPTION:
#    This script exports an identified SharePoint Online OR an on-premises SharePoint Server 2013/2016 Managed 
#    Metadata Service Group into a structured XML file. 
#    The XML file can then be passed to other SharePoint 2013+ system and imported to replicate the 
#    taxonomy structure in other environments by use of "ImportMMSGroup.PS1".
#    The key is that the taxonomy is replicated and should look the same in the replicated system meaning that
#    content can move between environments seamless from a managed metadata point of view.
#    This file can be run from SPO or from SharePoint On-Premises. If on-premises, then the user must be
#    specified as "DOMAIN\USER". SPO is detected by SHAREPOINT.COM being in the Admin URL.
#
#    Please note that this script was originally sourced from the community:
#      1) https://cann0nf0dder.wordpress.com/2014/11/12/exporting-taxonomy-from-sharepoint-using-powershell/
#      2) https://cann0nf0dder.wordpress.com/2014/11/29/importing-taxonomy-to-sharepoint-using-powershell/
#
#  PARAMETERS
#    $AdminUser - The user who has adminitrative access to the term store. (e.g., On-Premises: Domain\user Office365:user@<domain>.onmicrosoft.com)
#    $AdminPassword - The password for the Admin User "AdminUser".
#    $AdminUrl - The URL of Central Admin for on-premises or Admin site for Office 365
#    $PathToExportXMLTerms - The path to save the XML Output to. This path must exist.
#    $PathToSPClientdlls - The script requires the following CSOM dlls: (e.g. Office 365 - "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI" or on-premises - "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI")
#                            Microsoft.SharePoint.Client.dll
#                            Microsoft.SharePoint.Client.Runtime.dll
#                            Microsoft.SharePoint.Client.Taxonomy.dll
#    $GroupToExport - [string] Optional. If included only the identified Group will be exported. If omitted then the 
#                     entire termstore will be written to XML. This includes any site collections that have 
#                     their own termsets
#    $TermSetDumpIDs - [boolean] Optional. If set to $true, creates a <yyyymmdd>-<termset>.TXT in the path with name, description, MMS path and GUIDs
#
#  EXAMPLES
#    This exports the entire termstore of an Office 365 MMS with a domian of "domain"
#    ./ExportMMS.ps1 -AdminUser user@<tenant>.onmicrosoft.com -AdminPassword pass@word1 -AdminUrl https://<tenant>-admin.sharepoint.com -PathToExportXMLTerms c:\myTerms
#
#    This exports just the Term Store Group 'Subject Category (UK Defence Taxonomy)' and creates a termset dump file for reference
#    ./ExportMMS.ps1 -AdminUser user@<tenant>.onmicrosoft.com -AdminPassword pass@word1 -AdminUrl https://<tenant>-admin.sharepoint.com -PathToExportXMLTerms c:\myTerms -PathToSPClientdlls "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI" -GroupToExport 'Subject Category (UK Defence Taxonomy)' -TermSetDumpIDs $true
# 
#  VERSION HISTORY
#   DATE      WHO             VERSION  DESCRIPTION
#   23/05/16  Nigel Bridport  V1.0     Created initial file
#   13/07/16  Nigel Bridport  V1.1     New functionality. Addition of a termset dump option -TermSetDumpIDs
#                                      If set to True, then creates a TXT file in the filepath for the TermSet
#   12/10/16  Nigel Bridport  V1.2     Changes to the schema of the export so that it matches the transform schema
#                                      This enables the Merge process to function more correctly.
#
#  DISCLAIMER:
#  THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
#  MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES 
#  OF MERCHANTABILITY OR OF FITNESS FOR A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR 
#  PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR 
#  ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS, BUSINESS 
#  INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR 
#  INABILITY TO USE THE SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. 
#  BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION OF LIABILITY FOR CONSEQUENTIAL OR 
#  INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.
#
###############################################################################################################
# 
Param(
    [Parameter(Mandatory = $false)]
    [string]$AdminUser,

    [Parameter(Mandatory = $false)]
    [string]$AdminPassword,

    [Parameter(Mandatory = $false)]
    [string]$AdminUrl,

    [Parameter(Mandatory = $false)]
    [string]$PathToExportXMLTerms,

    [Parameter(Mandatory = $false)]
    [string]$PathToSPClientdlls,

    [Parameter(Mandatory = $false)]
    [string]$GroupToExport,

    [Parameter(Mandatory = $false)]
    [boolean]$TermSetDumpIDs
)


#######################################################################################
#
# Connect-MappingFile
# This function is used to create the logging file
#
#######################################################################################
function Connect-MappingFile($sFilePath, $bCreate){
    #Create the output termset dump file with necessary headers
    $fsStream = [System.IO.StreamWriter] $sFilePath
    if ($bCreate) {$fsStream.WriteLine("Short Name~Long Name~Path~GUID~Path|GUID")}

    return $fsStream
}


#####################################################################################################
#
# Get-XMLTermStoreTemplateToFile
# Creates a temporary template file that is used to populate from  information from the connected MMS. 
# Note: Nodes are just "replaced" in the code further down.
#
#####################################################################################################
function Get-XMLTermStoreTemplateToFile($termStoreName, $termStoreSystemGroup, $sFullFilePath){
    ## Set up an xml template used for creating your exported xml
    $xmlTemplate = '<TermStores>
        <TermStore Name="' + $termStoreName + '" IsOnline="True" WorkingLanguage="1033" DefaultLanguage="1033" SystemGroup="' + $termStoreSystemGroup + '">
            <Groups>
                <Group Id="" Name="" Description="" IsSystemGroup="False" IsSiteCollectionGroup="False">
                    <TermSets>
						<TermSet Id="" Name="" Description="" Contact="" IsAvailableForTagging="" IsOpenForTermCreation="" CustomSortOrder="" TermCount="">
                            <CustomProperties>
                                <CustomProperty Key="" Value=""/>
                            </CustomProperties>
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
        $xmlTemplate | Out-File($sFullFilePath)
 
        #Load file and return
        $xml = New-Object XML
        $xml.Load($sFullFilePath)
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
# The templates nodes are loaded for replacement as we enumerate the taxonomy
#
#####################################################################################################
function Get-XMLFileObjectTemplates($xml){
    $global:xmlGroupT = $xml.selectSingleNode('//Group[@Id=""]')  
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
# Ensures that the templates is clean and empty before starting to enumerate the 
# taxonomy with live data
#
#####################################################################################################
function Clean-Template($xml) {
    #Do not cleanup empty description nodes (this is the default state)

    ## Empty Term.Labels.Label
    $xml.selectnodes('//Label[@Value=""]') | ForEach-Object {
        $parent = $_.get_ParentNode()
        $parent.RemoveChild($_)  | Out-Null      
    } 
    
    ## Empty Term
    $xml.selectnodes('//Term[@Id=""]') | ForEach-Object {
        $parent = $_.get_ParentNode()
        $parent.RemoveChild($_)  | Out-Null      
    } 
    
    ## Empty TermSet
    $xml.selectnodes('//TermSet[@Id=""]') | ForEach-Object {
        $parent = $_.get_ParentNode()
        $parent.RemoveChild($_)  | Out-Null      
    } 
    
    ## Empty Group
    $xml.selectnodes('//Group[@Id=""]') | ForEach-Object {
        $parent = $_.get_ParentNode()
        $parent.RemoveChild($_)   | Out-Null     
    }
    
    ## Empty Custom Properties
    $xml.selectnodes('//CustomProperty[@Key=""]') | ForEach-Object {
        $parent = $_.get_ParentNode()
        $parent.RemoveChild($_) | Out-Null
    }

    ## Empty Local Custom proeprties
    $xml.selectnodes('//LocalCustomProperty[@Key=""]') | ForEach-Object {
        $parent = $_.get_ParentNode()
        $parent.RemoveChild($_) | Out-Null
    }

    $xml.selectnodes('//Descriptions')| ForEach-Object {
        $childNodes = $_.ChildNodes.Count
        
        if ($childNodes -gt 1)
        {
            $_.RemoveChild($_.ChildNodes[0]) | Out-Null
        }
    }

    while ($xml.selectnodes('//Term[@Id=""]').Count -gt 0)
    {
        #Cleanup the XML, remove empty Term Nodes
        $xml.selectnodes('//Term[@Id=""]').RemoveAll() | Out-Null
    }
}


#####################################################################################################
#
# ConnectToSP
# Loads the CSOM binaries and makes the initial connection to SharePoint Admin portal
#
#####################################################################################################
function ConnectToSP($url, $user, $securePassword){

    $spContext = New-Object Microsoft.SharePoint.Client.ClientContext($url)

    if($url.Contains(".sharepoint.com")) #SharePoint Online
    {	
	    $spCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($user, $securePassword)
    }
    else #SharePoint On-premises
    {	
	    $networkCredentials = New-Object -TypeName System.Net.NetworkCredential
	    $networkCredentials.UserName = $user.Split('\')[1]
	    $networkCredentials.SecurePassword = $securePassword
	    $networkCredentials.Domain = $user.Split('\')[0]

	    [System.Net.CredentialCache]$spCredentials = New-Object -TypeName System.Net.CredentialCache
	    $uri = [System.Uri]$url
	    $spCredentials.Add($uri, "NTLM", $networkCredentials)
    }

    #See if we can establish a connection
    $spContext.Credentials = $spCredentials
    $spContext.RequestTimeOut = 5000 * 60 * 10

    try
    {
        $spContext.ExecuteQuery()
        Write-Host "Successfully established a connection to SharePoint at $Url" -ForegroundColor Green
    }
    catch
    {
        Write-Host "Not able to connect to SharePoint on $Url. Exception:" $_.Exception.Message -ForegroundColor Red
        exit 1
    }
    return $spContext
}


#####################################################################################################
#
# Get-TermStoreInfo
# Gets the currently connected MMS service that this context is connected to
#
#####################################################################################################
function Get-TermStoreInfo($spContext){
    $spTaxSession = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($spContext)
    $spTaxSession.UpdateCache();
    $spContext.Load($spTaxSession)

    try
    {
        $spContext.ExecuteQuery()
    }
    catch
    {
        Write-Host "Error while loading the Taxonomy Session " $_.Exception.Message -ForegroundColor Red 
        exit 1
    }

    if ($spTaxSession.TermStores.Count -eq 0)
    {
        Write-Host "The Taxonomy Service is offline or missing" -ForegroundColor Red
        exit 1
    }

    $termStores = $spTaxSession.TermStores
    $spContext.Load($termStores)

    try
    {
        $spContext.ExecuteQuery()
        $termStore = $termStores[0]
        $spContext.Load($termStore)
        $spContext.ExecuteQuery()
        Write-Host "Connected to TermStore: " -NoNewLine
        Write-Host $($termStore.Name) -NoNewLine -ForegroundColor Green 
        Write-Host " ID: " -NoNewline -ForegroundColor White
        Write-Host $($termStore.Id) -ForegroundColor Green
    }
    catch
    {
        Write-host "Error details while getting term store ID" $_.Exception.Message -ForegroundColor Red
        exit 1
    }
    return $termStore
}


#####################################################################################################
#
# Get-Groups
# Gets either the supplied group or the whole group collection
#
#####################################################################################################
function Get-Groups($spContext, $groups, $xml, $groupToExport){

    #Loop through all groups, ignoring system Groups
    $groups | Where-Object { $_.IsSystemGroup -eq $false} | ForEach-Object{
   
        #Check if we are getting groups or just group.
        if ($groupToExport -ne "")
        {
            if ($groupToExport -ne $_.name)
            {
                #Return acts like a continue in ForEach-Object
                return;
            }
        }
    
        #Add each group to export xml by cloning the template group, populating it and appending it
        $xmlNewGroup = $global:xmlGroupT.Clone()
        $xmlNewGroup.Name = $_.name
        $xmlNewGroup.id = $_.id.ToString()
        $xmlNewGroup.Description = $_.description
        $xml.TermStores.TermStore.Groups.AppendChild($xmlNewGroup) | Out-Null

        Write-Host "Adding Group " -NoNewline
        Write-Host $_.name -ForegroundColor Green

        $spContext.Load($_.TermSets)
        try
        {
            $spContext.ExecuteQuery()     
            Get-TermSets $spContext $xmlNewGroup $_.Termsets $xml
        }
        catch
        {
            Write-host "Error while loading TermSets for Group " $xmlNewGroup.Name " " $_.Exception.Message -ForegroundColor Red
        }
    }
}


#####################################################################################################
#
# Get-TermSets accesses the TermSets that reside in the supplied Group or All Groups and then
# enumerates the entries for writing to the XML output file.
#
#####################################################################################################
function Get-TermSets($spContext, $xmlnewGroup, $termSets, $xml){
 
    $global:glCurrentTotalTerm = 0

    $termSets | ForEach-Object {
        #Add each termset to the export xml
        $xmlNewSet = $global:xmlTermSetT.Clone()
        #Replace SharePoint ampersand with regular
        $xmlNewSet.Name = $_.Name.replace("＆", "&")

        #Reset the current item count
        $global:glCurrentTotalTerm = 0
   
        $xmlNewSet.Id = $_.Id.ToString()
   
        if ($_.CustomSortOrder -ne $null) 
        { 
            $xmlNewSet.CustomSortOrder = $_.CustomSortOrder.ToString()            
        }

        foreach($customprop in $_.CustomProperties.GetEnumerator())
        {
            ## Clone Term customProp node
            $xmlNewTermCustomProp = $global:xmlTermCustomPropertiesT.Clone()    

            $xmlNewTermCustomProp.Key = $($customProp.Key)
            $xmlNewTermCustomProp.Value = $($customProp.Value)
            $xmlNewSet.CustomProperties.AppendChild($xmlNewTermCustomProp) | Out-Null 
        }

        $xmlNewSet.Description = $_.Description.ToString()
        $xmlNewSet.Contact = $_.Contact.ToString()
        $xmlNewSet.IsOpenForTermCreation = $_.IsOpenForTermCreation.ToString()  
        $xmlNewSet.IsAvailableForTagging = $_.IsAvailableForTagging.ToString()
        $xmlNewSet.TermCount = "0"
        $xmlNewGroup.TermSets.AppendChild($xmlNewSet) | Out-Null

        Write-Host "Adding TermSet " -NoNewline
        Write-Host $_.name -ForegroundColor Green -NoNewline
        Write-Host " to Group " -NoNewline
        Write-Host $xmlNewGroup.Name -ForegroundColor Green

        $spContext.Load($_.Terms)
        
        try
        {
            $spContext.ExecuteQuery()
        }
        catch
        {
            Write-host "Error while loading Terms for TermSet " $_.name " " $_.Exception.Message -ForegroundColor Red
            exit 1
        }
        
        #Find how many terms we have to work on to let the user know how long it might take
        # We also write this value into the export file so the import knows how many to expect.
        $terms = $_.GetAllTerms()
        $spContext.Load($terms)
        $spContext.ExecuteQuery()
        $lTotalTerms = $terms.Count

        #Write the number of terms out to the TermSet node
        $xmlNewSet.TermCount = $lTotalTerms.ToString()
       
        Get-Terms $spContext $_.Terms $xml 0 0 $lTotalTerms
    }
}


#####################################################################################################
#
# Get-Terms gets the supplied term from the MMS, extracts the properties and writes the information
# out to the appropriate XML node term.
#
#####################################################################################################
function Get-Terms($spContext, $terms, $xml, $lCurrentTermLevel, $global:glCurrentTotalTerms, $lTotalTerms){
    #$lCurrentTermLevel tells us at what level this term sits. SharePoint MMS can only support 7 levels currently
    ++$lCurrentTermLevel

    #Terms could be either the original termset or parent term with children terms
    $terms | ForEach-Object{

        #lCurrentTotalTerms is only used to indicate to the user how far through we currently are
        ++$global:glCurrentTotalTerms

        #Create a new term xml Element
        $xmlNewTerm = $global:xmlTermT.Clone()
        #Replace SharePoint ampersand with regular
        $xmlNewTerm.Name = $_.Name.replace("＆", "&")
        $xmlNewTerm.id = $_.Id.ToString()
        $xmlNewTerm.IsAvailableForTagging = $_.IsAvailableForTagging.ToString()
        $xmlNewTerm.IsKeyword = $_.IsKeyword.ToString()
	    $xmlNewTerm.IsReused = $_.IsReused.ToString()
	    $xmlNewTerm.IsRoot = $_.IsRoot.ToString()
        $xmlNewTerm.IsSourceTerm = $_.IsSourceterm.ToString()
        $xmlNewTerm.IsDeprecated = $_.IsDeprecated.ToString()

        if ($_.CustomSortOrder -ne $null)
        {
            $xmlNewTerm.CustomSortOrder = $_.CustomSortOrder.ToString()  
        }

        #Custom Properties
        foreach ($customprop in $_.CustomProperties.GetEnumerator())
        {
            # Clone Term customProp node
            $xmlNewTermCustomProp = $global:xmlTermCustomPropertiesT.Clone()    
        
            $xmlNewTermCustomProp.Key = $($customProp.Key)
            $xmlNewTermCustomProp.Value = $($customProp.Value)
            $xmlNewTerm.CustomProperties.AppendChild($xmlNewTermCustomProp)  | Out-Null
        }

        #Local Properties
        foreach ($localProp in $_.LocalCustomProperties.GetEnumerator())
        {
            # Clone Term LocalProp node
            $xmlNewTermLocalCustomProp = $global:xmlTermLocalCustomPropertiesT.Clone()    

            $xmlNewTermLocalCustomProp.Key = $($localProp.Key)
            $xmlNewTermLocalCustomProp.Value = $($localProp.Value)
            $xmlNewTerm.LocalCustomProperties.AppendChild($xmlNewTermLocalCustomProp) | Out-Null

            #CVRObjectID Sync
            if ($xmlNewTermLocalCustomProp.Key -eq "CVRObjectID")
            {
                $xmlNewTerm.CVRObjectID = $xmlNewTermLocalCustomProp.Value
            }
        }

        if ($_.Description -ne "")
        {
            $xmlNewTermDescription = $global:xmlTermDescriptionT.Clone()    
            $xmlNewTermDescription.Value = $_.Description
            $xmlNewTerm.Descriptions.AppendChild($xmlNewTermDescription) |Out-Null
        }
    
        $spContext.Load($_.Labels)
        $spContext.Load($_.TermSet)
        $spContext.Load($_.Parent)
        $spContext.Load($_.Terms)

        try
        {
            $spContext.ExecuteQuery()

            foreach ($label in $_.Labels)
            {  
                #Clone Term Label node
                $xmlNewTermLabel = $global:xmlTermLabelT.Clone()
                $xmlNewTermLabel.Value = $label.Value.ToString()
                $xmlNewTermLabel.Language = $label.Language.ToString()
                $xmlNewTermLabel.IsDefaultForLanguage = $label.IsDefaultForLanguage.ToString()
                $xmlNewTerm.Labels.AppendChild($xmlNewTermLabel) | Out-Null
            }

            #Use this terms parent term or parent termset in the termstore to find it's matching parent
            #in the export xml
            if ($_.parent.Id -ne $null)
            {
                #Both guids are needed as a term can appear in multiple termsets 
                $parentGuid = $_.parent.Id.ToString() 
                $parentTermsetGuid = $_.Termset.Id.ToString() 
                #$_.Parent.Termset.Id.ToString()
            }
            else 
            {
                $parentGuid = $_.Termset.Id.ToString() 
            }

            #Get this terms parent in the xml       
            $parent = Get-TermByGuid $xml $parentGuid $parentTermsetGuid
            $parentGuid = $null

            #Append new Term to Parent
            $term = $parent.Terms.AppendChild($xmlNewTerm)

            #Extra processing to dump out the term details into a text file
            if ($TermSetDumpIDs)
            {
                #Get the term path
                $xmlCurrentTerm = $term
                $sTermPath = ""
  
                do
                {
                    if ((-not ($xmlCurrentTerm.Name -eq "Terms")) -and  (-not ($xmlCurrentTerm.Name -eq "TermSets")))
                    {
                        $sTermPath = ":" + $xmlCurrentTerm.Name + $sTermPath
                    }
                    $xmlCurrentTerm = $xmlCurrentTerm.ParentNode
                }
                until ($xmlCurrentTerm.Name -eq "TermSets")

                #We need to chop off the TermSet name from the string
                $sTermPath = $sTermPath.Substring($sTermPath.IndexOf(":", 1) + 1)
                $sTermDescription = $term.Descriptions.Description[1].Value

                try
                {
                    $fsStream.WriteLine($term.Name + "~" + $sTermDescription + "~" + $sTermPath + "~" + $term.Id + "~" + $sTermPath + "|" + $term.Id)
                }
                catch
                {
                    Write-Host "Error writing to $sMappingFileName due to a locking issue for" $term.Name -ForegroundColor Red
                }
                finally
                {
                   #This enables the code to continue even if the write to dump file fails
                }
            }

            Write-Host "($global:glCurrentTotalTerms of $lTotalTerms) " -NoNewline
            Write-Host "Adding Level $lCurrentTermLevel Term " -NoNewline
            Write-Host $_.name -ForegroundColor Green -NoNewline
            Write-Host " to Parent " -NoNewline
            Write-Host $parent.Name -ForegroundColor Green

            #If this term has child terms we need to loop through those
            if ($_.Terms.Count -gt 0)
            {
                #Recursively call itself
                Get-Terms $spContext $_.Terms $xml $lCurrentTermLevel $global:glCurrentTotalTerms $lTotalTerms   
            }
        }
        catch
        {
            Write-host "Error while loaded addition information for Term" $xmlNewTerm.Name " " $_.Exception.Message -ForegroundColor Red
            #Log and retry?
            continue
        }
        finally
        {
            #Carry on with the next item
        }
    }
}


#####################################################################################################
#
# Get-TermByGuid attempts to get a current nodes parent from its GUID value
#
#####################################################################################################
function Get-TermByGuid($xml, $guid, $parentTermsetGuid){
    if ($parentTermsetGuid) 
    {
        return  $xml.selectnodes('//Term[@Id="' + $guid + '"]')
    } 
    else 
    {
        return  $xml.selectnodes('//TermSet[@Id="' + $guid + '"]') 
    }
}


#####################################################################################################
#
# ExportTaxonomy
#
#####################################################################################################
function ExportTaxonomy($spContext, $termStore, $xml, $groupToExport, $sPath, $saveFileName, $sTemporaryTemplateFileName){
   
    $spContext.Load($termStore.Groups)
    try
    {
        $spContext.ExecuteQuery();
    }
    catch
    {
        Write-host "Error while loaded Groups from TermStore " $_.Exception.Message -ForegroundColor Red
        exit 1
    }

    Get-Groups $spContext $termStore.Groups $xml $groupToExport

    #Clean up empty tags/nodes
    Clean-Template $xml

    #Save file.
    try
    {
        $sIntermediateFile = $sTemporaryTemplateFileName.Replace("-Template.XML", "-Taxonomy.XML")
        $xml.Save($sPath + "\" + $sIntermediateFile)
   
        #Clean up empty <Term> unable to work out in Clean-Template.
        Get-Content ($sPath + "\" + $sIntermediateFile) | Foreach-Object { $_ -replace "<Term><\/Term>", "" } | Set-Content ($sPath + "\" + $saveFileName)
        Write-Host "Saving XML EXPORT file to [$sPath\$saveFileName]" -ForegroundColor Yellow

        #Remove temp file
        Remove-Item($sPath + "\" + $sTemporaryTemplateFileName);
        Remove-Item($sPath + "\" + $sIntermediateFile);
    }
    catch
    {
        Write-Host "Error saving XML File to disk " $_.Exception.Message -ForegroundColor Red
        exit 1
    }
}


#####################################################################################################
#
# This is where the code will run.
#
#####################################################################################################

#Inform the user how long the operation took.  This is the start date/time.
$dtStartTime = Get-Date
$sDate = (Get-Date).ToString("yyyyMMdd") #Used for any file naming

Write-Host "##################################################################" -ForegroundColor Cyan
Write-Host "Starting the MMS GROUP EXPORT at $dtStartTime" -ForegroundColor Cyan
Write-Host "##################################################################" -ForegroundColor Cyan

#Initialise any variables.
if ($AdminUser -eq "") {$AdminUser = Read-Host "What SharePoint Administrator user account, with Termstore Admin rights, do you want to use?"}
if ($AdminPassword.Length -eq 0) {[System.Security.SecureString]$AdminPassword = Read-Host -AsSecureString "What is the password for $AdminUser ?"}

if ($AdminUser.Contains("\"))
{
    $PathToSPClientdlls = "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI" #Default for SP on-premises       

    if ($AdminUrl -eq "") {$AdminUrl = Read-Host "What is the SharePoint Central Administration URL to export from?"} #If the user hasn't specified a URL, ask them for the tenant and fix-up

    $sHost = $AdminUrl.Substring($AdminUrl.IndexOf("//") + 2)
    $sHost = $sHost.Substring(0, $sHost.IndexOf(":"))
}
else
{
    $PathToSPClientdlls = "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI" #Default for SPO CSOM

    if ($AdminUrl -eq "")
    {
        $AdminUrl = Read-Host "Which TENANT do you want to EXPORT from?" #If the user hasn't specified a URL, ask them for the tenant and fix-up
        $AdminUrl = "https://$AdminUrl-admin.sharepoint.com"
    }

    $sHost = $AdminUrl.Substring($AdminUrl.IndexOf("//") + 2, $AdminUrl.IndexOf("-admin.") - $AdminUrl.IndexOf("//") - 2)
}

if (($GroupToExport -eq "") -or ($GroupToExport -eq $null)) {[string]$GroupToExport = Read-Host "You have selected to export all Groups from the MMS. This could be a large set. Press <RETURN> to confirm or enter the name of the GROUP to Export"} #User this to specify the MMS GROUP that you want to export. Granularity is at the GROUP level only - If left blank, ALL GROUPS will be exported, i.e. the whole MMS structure
if ($GroupToExport -eq "")
{
    $sFile = "ALLGroups"
}
else
{
    $sFile = $GroupToExport.Replace(" ", "").ToUpper()
}

#Use this to specify a path to save the export file to. If blank, will export to the current location - Can also control via the $PathToExportXMLTerms argument.
if ($PathToExportXMLTerms -eq "")
{
    $sPathTemp = Get-Location
    $PathToExportXMLTerms = $sPathTemp.Path
}

#Create the Terms Export filename
$XMLTermsFileName = ($sDate + "-" + $sFile + "-" + $sHost.ToUpper() + ".XML")

#Switch to enable the dump of a termset. This will pull all of the termset terms, minus structure, and output with their respective GUID values
if ($TermSetDumpIDs)
{
    Write-Host "Creating a dump file of terms and their associated GUID values..." -ForegroundColor Yellow
    $sMappingFileName = [string]$PathToExportXMLTerms + "\" + $sDate + "-" + $sFile + "-" + $sHost.ToUpper() + ".TXT"

    if (-not (Test-Path $sMappingFileName)) 
    {
        $fsStream = Connect-MappingFile $sMappingFileName $true
    }
    else
    {
        $fsStream = Connect-MappingFile $sMappingFileName $false
    }
}

#Test for existence and add client dlls references
try
{
    if (Test-Path "$PathToSPClientdlls\Microsoft.SharePoint.Client.dll")
    {
        Add-Type -Path "$PathToSPClientdlls\Microsoft.SharePoint.Client.dll"
        Add-Type -Path "$PathToSPClientdlls\Microsoft.SharePoint.Client.Runtime.dll"
        Add-Type -Path "$PathToSPClientdlls\Microsoft.SharePoint.Client.Taxonomy.dll"
        Add-Type -Path "$PathToSPClientdlls\Microsoft.SharePoint.Client.DocumentManagement.dll"

        #Connect to SharePoint Online - If the user was requested for a password, it will already be a securestring
        if ($AdminPassword.ToString() -ne "System.Security.SecureString") {[System.Security.SecureString]$AdminPassword = ConvertTo-SecureString $AdminPassword -AsPlainText -Force}

        $spContext = ConnectToSP $AdminUrl $AdminUser $AdminPassword
        $termStore = Get-TermStoreInfo $spContext

        #The 3 lines below just helps to create the template file with all of the correct detail.
        $termStoreSystemGroup = $termStore.SystemGroup
        $spContext.Load($termStoreSystemGroup)
        $spContext.ExecuteQuery()

        #Create the temport template XML filename
        $sTemplateFileName = (Get-Date).ToString("yyyyMMddHHmmss") + "-Template.XML" #Used for any file naming

        $xmlFile = Get-XMLTermStoreTemplateToFile $termStore.Name $termStoreSystemGroup.Id ($PathToExportXMLTerms + "\" + $sTemplateFileName)
        Get-XMLFileObjectTemplates $xmlFile

        ExportTaxonomy $spContext $termStore $xmlFile $GroupToExport $PathToExportXMLTerms $XMLTermsFileName $sTemplateFileName
    }
    else
    {
        Write-Host "The path $PathToSPClientdlls does not point to the SharePoint CSOM libraries. Please update and try again." -ForegroundColor Red
    }
}
catch
{

}
finally
{
    #Disposal of objects
    $spContext.Dispose()

    #If a dump file has been requested
    if ($TermSetDumpIDs) 
    {
        Write-Host "Saving TXT GUID dumpfile to [$sMappingFileName]" -ForegroundColor Yellow
        $fsStream.Flush()
        $fsStream.Close()
        $fsStream.Dispose()
    }

    #Work out the elapsed time.
    $dtEndTime = Get-Date
    $dtDifference = New-TimeSpan -Start $dtStartTime -End $dtEndTime

    Write-Host "##################################################################" -ForegroundColor Cyan
    Write-Host "Completed..." -ForegroundColor Cyan
    Write-Host "The Taxonomy export operation for the group $GroupToExport" -ForegroundColor Cyan 
    Write-Host "took $dtDifference" -ForegroundColor Cyan
    Write-Host "##################################################################" -ForegroundColor Cyan
}
#End of script