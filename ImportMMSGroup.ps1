###############################################################################################################
# 
#  NAME:  ImportMMSGroup.ps1
#
#  VERSION: V1.0
# 
#  DATE: 23rd May 2016
#
#  DESCRIPTION:
#    This script exports an identified SharePoint Online OR an on-premises SharePoint Server 2013/2016 Managed 
#    Metadata Service Group into a structured XML file. 
#    The XML file can then be passed to other SharePoint 2013+ system and imported to replicate the 
#    taxonomy structure in other environments by use of "ImportMMSGroup.PS1".
#    The key is that the taxonomy is replicated and should look the same in the replicated system meaning that
#    content can move between environments seamless from a managed metadata point of view.
#
#    Please note that this script was originally sourced from the community:
#      1) https://cann0nf0dder.wordpress.com/2014/11/12/exporting-taxonomy-from-sharepoint-using-powershell/
#      2) https://cann0nf0dder.wordpress.com/2014/11/29/importing-taxonomy-to-sharepoint-using-powershell/
#
#  PARAMETERS
#    $AdminUser - The user who has adminitrative access to the term store. (e.g., On-Premises: Domain\user Office365:user@<domain>.onmicrosoft.com)
#    $AdminPassword - The password for the Admin User "AdminUser".
#    $AdminUrl - The URL of Central Admin for on-premises or Admin site for Office 365
#    $TermsFilePath - The path to read the XML Output from.
#    $XMLTermsFileName - The name of the XML file to save. If the file already exists then it will be overwritten.
#    $Update - $false or $true: Controls the speed of updates. Set to $false, the default, the import will pull each term live from SPO and compare 
#              with the XML vaersion to note change. The will enumerate and check every term. If you are looking to update the termset with changes
#              from an external tool, i.e. RDM, and have compared that output already with an export, via MergeXMLs.PS1, then switch this to $true.
#              The import process will then only look to work with terms that it has to and expect other terms to be correct. Greatly improves performance.
#
#  EXAMPLES
#    This imports the supplied XML to an Office 365 MMS with a domian of "<tenant>"
#    ./ImportMMSGroup.ps1 -AdminUser user@<tenant>.onmicrosoft.com -AdminPassword pass@word1 -AdminUrl https://<tenant>-admin.sharepoint.com -TermsFilePath c:\myTerms\20160523-MMSExportedterms.xml -PathToSPClientdlls "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI"
#
#    This imports just the Term Store Group from the presented XML
#    ./ImportMMSGroup.ps1 -AdminUser user@<tenant>.onmicrosoft.com -AdminPassword pass@word1 -AdminUrl https://<tenant>-admin.sharepoint.com -TermsFilePath c:\myTerms\20160523-SubjectCategoryExportedterms.xml
#
#  VERSION HISTORY
#   DATE      WHO             DESCRIPTION
#   23/05/16  Nigel Bridport  Created initial file
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
    [string]$TermsFilePath,

    [Parameter(Mandatory = $false)]
    [boolean]$Update
)


#####################################################################################################
#
# Get-TermStoreInfo
# Gets the currently connected MMS service that this context is connected to
#
#####################################################################################################
function Get-TermStoreInfo($spContext){
    $spTaxSession = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($spContext)
    $spTaxSession.UpdateCache()
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
        write-host "The Taxonomy Service is offline or missing" -ForegroundColor Red
        exit 1
    }

    $mmsTermStores = $spTaxSession.TermStores
    $spContext.Load($mmsTermStores)

    try
    {
        #This tries to get the first instance of the MMS service application. Always valid for Office 365
        # environments but may not be so for on-premises. If on-premises, then assumes that the termset is 
        # in the first MMS service applciation as there could potentially be more.
        $spContext.ExecuteQuery()
        $mmsTermStore = $mmsTermStores[0]
        $spcontext.Load($mmsTermStore)
        Write-Host "Connected to TermStore: $($mmsTermStore.Name) ID: $($mmsTermStore.Id)"
    }
    catch
    {
        Write-host "Error details while getting term store ID" $_.Exception.Message -ForegroundColor Red
        exit 1
    }
    return $mmsTermStore
}


#####################################################################################################
#
# Get-TermsToImport
# Load the XML file into memory
#
#####################################################################################################
function Get-TermsToImport($xmlTermsPath){
    [Reflection.Assembly]::LoadWithPartialName("System.Xml.Linq") | Out-Null

    try
    {
        $xmlDocument = [System.Xml.Linq.XDocument]::Load($xmlTermsPath, [System.Xml.Linq.LoadOptions]::None)
        return $xmlDocument
    }
    catch
    {
        Write-Host "Unable to read ExportedTermsXML. Exception:$_.Exception.Message" -ForegroundColor Red
        exit 1
    }
}


#####################################################################################################
#
# Check-Groups
# Check the specified MMS Groups as defined by the import XML
#
#####################################################################################################
function Check-Groups($spContext, $mmsTermStore, $xmlTerms){
     
     foreach ($xmlGroupNode in $xmlTerms.Descendants("Group"))
     {
        $xmlGroupName = $xmlGroupNode.Attribute("Name").Value
        $xmlGroupId = $xmlGroupNode.Attribute("Id").Value

        #Check to make sure we are working with a valid file.
        # If "Id='TRANSFORMED'", then this is the Src processed file. It has to be "Merged" with a Live Export before attempting to re-import
        if ($xmlGroupId.ToUpper() -eq "TRANSFORMED")
        {
            Write-Host "INCORRECT INPUT FILE PRESENTED" -ForegroundColor Red
            Write-Host "The [-TermsFilePath] argument passed in was [" -NoNewList -ForeGroundColor Cyan
            Write-Host $TermsFilePath -NoNewLine -ForegroundColor White
            Write-Host "]" -ForegroundColor Cyan 
            Write-Host "This file is a TRANSFORMED Version of the Source (not a SPO MMS Export)" -NoNewline -ForeGroundColor Cyan
            Write-Host "Please use MergeXMLs.ps1 with $TermsFilePath as the -SRCTerm argument such as:" -ForeGroundColor Cyan
            Write-Host "> MergeXMLs.ps1 -FilePath <directory where both of the folling files must exist> -SRCExport '$TermsFilePath' -LIVEExport '<filename of the export file from SPO MMS>'" -ForGroundColor White
        }
        else
        {
            $xmlGroupGuid = [System.Guid]::Parse($xmlGroupId)
            Write-Host "Processing Group: " -NoNewline -ForegroundColor White
            Write-Host $xmlGroupName -NoNewline -ForegroundColor Green
            Write-Host " ID: " -NoNewLine -ForegroundColor White
            Write-Host $xmlGroupId -NoNewline -ForegroundColor Green 
            Write-Host " ... " -NoNewline -ForegroundColor White

            $xmlGroup = $mmsTermStore.GetGroup($xmlGroupGuid)
            $spContext.Load($xmlGroup)
        
            try
            {
                $spContext.ExecuteQuery()
            }
            catch
            {
                Write-Host "Error while finding if $xmlGroupName group already exists. " $_.Exception.Message -ForegroundColor Red 
                exit 1
            }
	
	        if ($xmlGroup.ServerObjectIsNull)
            {
                $xmlGroup = $mmsTermStore.CreateGroup($xmlGroupName, $xmlGroupGuid);
                $spContext.Load($xmlGroup);
                try
                {
                    $spContext.ExecuteQuery();
		            Write-Host "inserted" -ForegroundColor Green
                }
                catch
                {
                    Write-Host "Error creating new Group $xmlGroupName. " $_.Exception.Message -ForegroundColor Red 
                    exit 1
                }
            }
    	    else
            {
		        Write-Host "already exists" -ForegroundColor White
	        }
	        Check-TermSets $xmlTerms $xmlGroup $mmsTermStore $spContext
        }

        try
        {
            $mmsTermStore.CommitAll()
            $spContext.ExecuteQuery()
        }
        catch
        {
            Write-Host "Error commiting changes to server. Exception:$_.Exception.Message" -ForegroundColor Red
            exit 1
        }
    }
}


#####################################################################################################
#
# Check-TermSets
# Check the necessary TermSet object from the import XML file
#
#####################################################################################################
function Check-TermSets($xmlTerms, $xmlGroup, $mmsTermStore, $spContext) {
	
    #Process the termsets nodes
    $termSets = $xmlTerms.Descendants("TermSet") | Where { $_.Parent.Parent.Attribute("Name").Value -eq $xmlGroup.Name }

	foreach ($termSetNode in $termSets)
    {
        $errorOccurred = $false

		$tsName = $termSetNode.Attribute("Name").Value
        $tsId = [System.Guid]::Parse($termSetNode.Attribute("Id").Value)
        $tsDescription = $termSetNode.Attribute("Description").Value
        $tsCustomSortOrder = $termSetNode.Attribute("CustomSortOrder").Value
        $global:glTermsetTermCount = $termSetNode.Attribute("TermCount").Value #Total number of terms in the original set
        Write-Host "Processing TermSet " -NoNewline -ForegroundColor White
        Write-Host $tsName -NoNewline -ForegroundColor Green
        Write-Host " ... " -NoNewLine -ForegroundColor White
		
		$termSet = $mmsTermStore.GetTermSet($tsId)
        $spcontext.Load($termSet)
                
        try
        {
            $spContext.ExecuteQuery()
        }
        catch
        {
            Write-Host "Error while finding if $tsName termset already exists. " $_.Exception.Message -ForegroundColor Red 
            exit 1
        }

		if ($termSet.ServerObjectIsNull)
        {
            #Termset could not be found. Creating it with the correct properties
			$termSet = $xmlGroup.CreateTermSet($tsName, $tsId, $mmsTermStore.DefaultLanguage)
            $termSet.Description = $tsDescription

            if ($tsCustomSortOrder -ne "")
            {
                $termSet.CustomSortOrder = $tsCustomSortOrder
            }

            $termSet.IsAvailableForTagging = [bool]::Parse($termSetNode.Attribute("IsAvailableForTagging").Value)
            $termSet.IsOpenForTermCreation = [bool]::Parse($termSetNode.Attribute("IsOpenForTermCreation").Value)

            if ($termSetNode.Element("CustomProperties") -ne $null)
            {
                foreach ($custProp in $termSetNode.Element("CustomProperties").Elements("CustomProperty"))
                {
                    $termSet.SetCustomProperty($custProp.Attribute("Key").Value, $custProp.Attribute("Value").Value)
                }
            }

            try
            {
                $spContext.ExecuteQuery()
            }
            catch
            {
                Write-Host "Error occured while creating the termset" $tsName $_.Exception.Message -ForegroundColor Red
                $errorOccurred = $true
            }
            Write-Host "created" -ForegroundColor Yellow
		}
		else
        {
            #The correct termset exists
			Write-Host "already exists" -ForegroundColor White
		}

        $global:glTermCurrentCount = 0 #Just a current term counter

        if (!$errorOccurred)
        {
            if ($termSetNode.Element("Terms") -ne $null)
            {
                foreach ($termNode in $termSetNode.Element("Terms").Elements("Term"))
                {
                    Check-Term $termNode $true $termSet $mmsTermStore $mmsTermStore.DefaultLanguage $spContext $global:glTermCurrentCount
                }
            }
            $termSetNode.Document.Save($TermsFilePath)
        }


            #Fixup actions.
        if ($termSetNode.Element("InvalidTerms") -ne $null)
        {
            Write-Host ""
            Write-Host "All existing terms have been processed. Now to look for fixup actions." -ForegroundColor White
            Write-Host "These actions will either [Deprecate], [Remove] or [Do Nothing] to the terms from the SOURCE export file" -ForegroundColor White
            Write-Host "as defined at the time of the merge process." -ForegroundColor White
            Write-Host "Terms being managed are:" -ForegroundColor White
        
            $lTotalCount = 0 #Totals counter
           
            foreach ($termNode in $termSetNode.Element("InvalidTerms").Elements("Fixup"))
            {
                Write-Host $termNode.Attribute("Action").Value -NoNewline -ForegroundColor Green
                Write-Host " | " -NoNewline -ForegroundColor White
                Write-Host $termNode.Attribute("Name").Value -ForegroundColor Green
                ++$lTotalCount
            }

            if ($lTotalCount -eq 0)
            {
                Write-Host ""
                Write-Host "There are no defined Fixup actions for this termset." -ForegroundColor White
                Write-Host "Termset is up-to-date" -ForegroundColor White
                Write-Host ""
            }
            else
            {
                $lCurrentCount = 1 #Current item counter

                #Check with the user before executing the command
                foreach ($termNode in $termSetNode.Element("InvalidTerms").Elements("Fixup"))
                {
                    Fixup-Term $termNode $termSet $mmsTermStore $spContext $lCurrentCount $lTotalCount
                    ++$lCurrentCount
                }
            }

            $termSetNode.Document.Save($TermsFilePath)
            Write-Host ""
        }
        else
        {
            Write-Host ""
            Write-Host "There are no defined Fixup actions for this termset." -ForegroundColor White
            Write-Host "Termset is up-to-date" -ForegroundColor White
            Write-Host ""
        }
    }
}


#####################################################################################################
#
# Fixup-Term
# This function fixes up the term from the supplied content. This effects terms that need to be
# deleted, or deprecated.
# They are hosted within the import file within the <InvalidTerms/> node.
#
#####################################################################################################
function Fixup-Term($xmlTerm, $mmsTermSet, $mmsStore, $spContext, $lCurrentCount, $lTotalCount){

    Write-Host "($lCurrentCount of $lTotalCount) Processing Term " -NoNewline -ForegroundColor White
    Write-Host $xmlTerm.Attribute("Name").Value -NoNewline -ForegroundColor Green
    Write-Host " ... " -NoNewline -ForegroundColor White

    #Get the term
    $mmsTerm = $termSet.GetTerm([System.Guid]::Parse($xmlTerm.Attribute("Id").Value))
    $spContext.Load($mmsTerm)
    $spContext.ExecuteQuery()

    if ($mmsTerm -ne $null)
    {
        #Perform the action as defined in the XML file
        try
        {
            switch ($xmlTerm.Attribute("Action").Value.ToLower())
            {
                "deprecate"
                {
                    $mmsTerm.Deprecate($true) | Out-Null
                    $spContext.ExecuteQuery()
                    Write-Host "deprecated" -ForegroundColor Yellow
                }

                "remove" 
                {
                    $mmsTerm.DeleteObject() | Out-Null
                    $spContext.ExecuteQuery()
                    Write-Host "deleted" -ForegroundColor Yellow
                }

                "nothing"
                {
                    Write-Host "no action" -ForegroundColor White
                }
            }
            $xmlTerm.Attribute("Processed").Value = $true
        }
        catch
        {
            Write-Host "ERROR "$_.Exception.Message -ForegroundColor Red
        }
    }
    else
    {
        Write-Host
    }
}


#####################################################################################################
#
# Check-Term
# This function creates the term from the supplied content
#
#####################################################################################################
function Check-Term($xmlTerm, $bRootNode, $mmsTermSet, $mmsStore, $lcid, $spContext, $global:glTermCurrentCount){

    ++$global:glTermCurrentCount #Increment the current term counter
    
    Write-Host "($global:glTermCurrentCount of $global:glTermsetTermCount) Processing Term " -NoNewline
    Write-Host $xmlTerm.Attribute("Name").Value -NoNewline -ForegroundColor Green
    Write-Host " ... " -NoNewLine -ForegroundColor White

    #There maybe problems with timeout issues or connection closures.
    # We set the "Processed" attribute for each term when it has been actioned
    # Check that attribute before doing anywork on the current term
    $bLiveCheck = $true

    #$Update is an input argument. If set to $true, it assumes that the export is a copy of the current live termset
    # so only action the terms that have been modified. If $false, it will get every term from the termstore, compare with 
    # the XML and then decide whether to update or not.
    # This can really speed things up for those large termsets.
    if ($Update)
    {
        if (($xmlTerm.Attribute("Processed").Value -eq "No Changes") -or ($xmlTerm.Attribute("Processed").Value -eq "False"))
        {
            $bLiveCheck = $false
        }
    }
    elseif ($xmlTerm.Attribute("Processed").Value -eq $true)
    {
        $bLiveCheck = $false
    }

    #Decide whether to do a live check of the term or trust the export/processed XML file.
    if ($bLiveCheck)
    {    
        $xmlTermName = $xmlTerm.Attribute("Name").Value

        #Get the source ID (GUID). If it is nothing "", then this is a new term and requires a new GUID
        if ($xmlTerm.Attribute("Id").Value.Length -eq 0)
        {
            $xmlTermId = [System.Guid]::NewGuid().ToString()
            $xmlTerm.Attribute("Id").Value = $xmlTermId #Update the XML with the value as it will be needed later. Particularly, if it has child terms or terms that need to move beneath it
        }
        else
        {
            $xmlTermId = [System.Guid]::Parse($xmlTerm.Attribute("Id").Value)
        }

        #Get the source Sort Order.
        if (($xmlTerm.Attribute("CustomerSortOrder").Value -ne $null) -or ($xmlTerm.Attribute("CustomerSortOrder").Value -ne ""))
        {
            $xmlTermCustomSortOrder = $xmlTerm.Attribute("CustomSortOrder").Value
        }
        else
        {
            $xmlTermCustomSortOrder = ""
        }

        #Get the source IsReused value
        if ($xmlTerm.Attribute("IsReused").Value.Length -eq 0)
        {
            $xmlTermIsReused = $false
        }
        else
        {
            $xmlTermIsReused = [bool]::Parse($xmlTerm.Attribute("IsReused").Value)
        }

        #Get the source IsSourceTerm - This is important for pinned/reused terms
        if ($xmlTerm.Attribute("IsSourceTerm").Value.Length -eq 0)
        {
            $xmlTermIsSourceTerm = $false
        }
        else
        {
            $xmlTermIsSourceTerm = [bool]::Parse($xmlTerm.Attribute("IsSourceTerm").Value)
        }

        #Get the source IsRoot value. This is important for pinned/reused terms
        if ($xmlTerm.Attribute("IsRoot").Value.Length -eq 0)
        {
            $xmlTermIsRoot = $false
        }
        else
        {
            $xmlTermIsRoot = [bool]::Parse($xmlTerm.Attribute("IsRoot").Value)
        }

        #Get the sources ability for tagging
        if ($xmlTerm.Attribute("IsAvailableForTagging").Value.Length -eq 0)
        {
            $xmlTermIsAvailableForTagging = $true
        }
        else
        {
            $xmlTermIsAvailableForTagging = [bool]::Parse($xmlTerm.Attribute("IsAvailableForTagging").Value)
        }

        $mmsSourceTerm = $null
        $errorOccurred = $false

        if ($xmlTermIsReused)
        {
            if (!$xmlTermIsRoot -and !$xmlTermIsSourceTerm)
            {
                $xmlSourceTermSetId = [System.Guid]::Parse($xmlTerm.Attribute("SourceTermSetId").Value)
                $xmlSourceTermId =  [System.Guid]::Parse($xmlTerm.Attribute("SourceTermId").Value)
                $mmsSourceTerm = $mmsStore.GetTermInTermSet($xmlSourceTermSetId, $xmlSourceTermId)
                $spContext.Load($mmsSourceTerm)
            } 
        }

        #Get the term and its current parent
        $mmsTerm = $mmsTermSet.GetTerm($xmlTermId)
        $spContext.Load($mmsTerm)

        try
        {
            $spContext.ExecuteQuery()
        
            #If we have a live term, try to get its parent
            if (-not $mmsTerm.ServerObjectIsNull)
            {
                if (-not $bRootNode)
                {
                    $mmsParentTerm = $mmsTerm.Parent
                    $spContext.Load($mmsParentTerm)
                    $spContext.ExecuteQuery()
                }
            }
            else
            {
                #This is a term that needs creating. We need the parent node first. 
                #The id of the parent node should be in the XML source, even if it has just been created as we keep the info up-to-date
                if (-not $bRootNode)
                {
                    $mmsParentTerm = $mmsTermSet.GetTerm($xmlTerm.Parent.Parent.Attribute("Id").Value)
                    $spContext.Load($mmsParentTerm)
                    $spContext.ExecuteQuery()
                }
            }
        }
        catch
        {
            Write-Host "Error while finding if $xmlTermName term id already exists. " $_.Exception.Message -ForegroundColor Red
            
            #Save the status
            $xmlTerm.Document.Save($TermsFilePath)
        }
    
        if (($mmsTerm.ServerObjectIsNull) -and ($xmlTermIsSourceTerm))
        {
            #This is a term that needs to be created
            if ($mmsSourceTerm -ne $null)
            {
                if ($mmsParentTerm -ne $null) 
                {
                    $mmsTerm = $mmsParentTerm.reuseTerm($mmsSourceTerm, $false)
                }
                else 
                {
                    $mmsTerm = $mmsTermSet.reuseTerm($mmsSourceTerm, $false)
                }
            }
            elseif ($mmsParentTerm -ne $null)
            {
                $mmsTerm = $mmsParentTerm.CreateTerm($xmlTermName, $lcid, $xmlTermId)
            }
            else 
            {
                $mmsTerm = $mmsTermSet.CreateTerm($xmlTermName, $lcid, $xmlTermId)
            }

            $mmsTerm.IsAvailableForTagging = $xmlTermIsAvailableForTagging
        
            if ($xmlTermCustomSortOrder -ne "")
            {
                $mmsTerm.CustomSortOrder = $xmlTermCustomSortOrder
            }

            #Set the deprecated value
            if ([bool]::Parse($xmlTerm.Attribute("IsDeprecated").Value))
            {
                $mmsTerm.Deprecate($true)
            }
            else 
            {
                $mmsTerm.Deprecate($false)
            }

            if ($xmlTerm.Element("LocalCustomProperties") -ne $null)
            {
                foreach ($xmlLocalCustProp in $xmlTerm.Element("LocalCustomProperties").Elements("LocalCustomProperty"))
                {
                    $mmsTerm.SetLocalCustomProperty($xmlLocalCustProp.Attribute("Key").Value, $xmlLocalCustProp.Attribute("Value").Value)
                }
            }
        
            if ($xmlTerm.Element("Labels") -ne $null)
            {
                foreach ($xmlLabel in $xmlTerm.Element("Labels").Elements("Label"))
                {
                    #We ignore the first True Label as this is the default label.
                    if ([bool]::Parse($xmlLabel.Attribute("IsDefaultForLanguage").Value) -ne $true)
                    {
                        $mmsLabelTerm = $mmsTerm.CreateLabel($xmlLabel.Attribute("Value").Value, [int]$xmlLabel.Attribute("Language").Value, [bool]::Parse($xmlLabel.Attribute("IsDefaultForLanguage").Value))
                    }
                }
            }

            #Only update if not reused term.
            if ($mmsSourceTerm -eq $null)
            {
                $xmlDescription = $xmlTerm.Element("Descriptions").Element("Description").Attribute("Value").Value
                $mmsTerm.SetDescription($xmlDescription, $lcid)
        
                if ($xmlTerm.Element("CustomProperties") -ne $null)
                {
                    foreach ($xmlCustProp in $xmlTerm.Element("CustomProperties").Elements("CustomProperty"))
                    {
                        $mmsTerm.SetCustomProperty($xmlCustProp.Attribute("Key").Value, $xmlCustProp.Attribute("Value").Value)
                    }
                }
            }

            try
            {
                $spContext.Load($mmsTerm);
                $spContext.ExecuteQuery();
	            Write-Host "created" -ForegroundColor Yellow

                #Update the XML Document "Processed" attribute in case we need to rerun the actions
                $xmlTerm.Attribute("Processed").Value = $true
	        }
            catch
            {
                Write-Host "Error occured while creating term" $xmlTermName $_.Exception.Message -ForegroundColor Red
                $errorOccurred = $true

                #Save the status
                $xmlTerm.Document.Save($TermsFilePath)
            }
        }
        elseif ($xmlTermIsSourceTerm)
        {
            #This may be a term modification. Check the requested config with what is currently live
            $bTermModified = $false

            #There is an odd behaviour with the & and how SP encodes it. Just check and replace.
            #$mmsTerm.Name = $mmsTerm.Name.replace("＆", "&")

            #Check the name
            if ($mmsTerm.Name.CompareTo($xmlTerm.Attribute("Name").Value) -ne 0)
            {
                $mmsTerm.Name = $xmlTerm.Attribute("Name").Value
                $bTermModified = $true
            }

            if ($mmsTerm.IsAvailableForTagging -ne [bool]::Parse($xmlTerm.Attribute("IsAvailableForTagging").Value))
            {
                $mmsTerm.IsAvailableForTagging = [bool]::Parse($xmlTerm.Attribute("IsAvailableForTagging").Value)
                $bTermModified = $true
            }

            #There is an issue with "" = $null from XML. Initialise first.
            If ($mmsTerm.CustomSortOrder -eq $null) {$mmsTerm.CustomSortOrder = ""}
        
            if ($mmsTerm.CustomSortOrder -ne $xmlTerm.Attribute("CustomSortOrder").Value)
            {
                $mmsTerm.CustomSortOrder = $xmlTerm.Attribute("CustomSortOrder").Value
                $bTermModified = $true
            }
        
            if ($mmsTerm.IsDeprecated -ne [bool]::Parse($xmlTerm.Attribute("IsDeprecated").Value))
            {
                $mmsTerm.Deprecate([bool]::Parse($xmlTerm.Attribute("IsDeprecated").Value))
                $bTermModified = $true
            }
        
            if ($mmsTerm.IsKeyword -ne [bool]::Parse($xmlTerm.Attribute("IsKeyword").Value))
            {
                $mmsTerm.IsKeyword = [bool]::Parse($xmlTerm.Attribute("IsKeyword").Value)
                $bTermModified = $true
            }
        
            if ($mmsTerm.IsReused -ne [bool]::Parse($xmlTerm.Attribute("IsReused").Value))
            {
                #Only valid to set the ReUsed property if it is not a SourceTerm
                if (-not $xmlTermIsSourceTerm)
                {
                    $mmsTerm.IsReused = [bool]::Parse($xmlTerm.Attribute("IsReused").Value)
                    $bTermModified = $true
                }
            }
        
            if ($mmsTerm.IsSourceTerm -ne [bool]::Parse($xmlTerm.Attribute("IsSourceTerm").Value)) 
            {
                $mmsTerm.IsSourceTerm = [bool]::Parse($xmlTerm.Attribute("IsSourceTerm").Value)
                $bTermModified = $true
            }
        
            if ($mmsTerm.Description -ne $xmlTerm.Element("Descriptions").Element("Description").Attribute("Value").Value)
            {
                $mmsTerm.SetDescription($xmlTerm.Element("Descriptions").Element("Description").Attribute("Value").Value, $lcid)
                $bTermModified = $true
            }
        


            #Check the term labels
            # Read out all of the SPO MMS term labels to check against later
            $sTermLabels = "|"
            $mmsLabels = $mmsTerm.Labels
            $spContext.Load($mmsLabels)
            $spContext.ExecuteQuery()

            for ($iCount = 1; $iCount -lt $mmsLabels.Count; $iCount++)
            {
                $sTermLabels += $mmsLabels[$iCount].Value + "|"
            }

            #Now compare and check the XML and Term labels
            if ($xmlTerm.Element("Labels") -ne $null)
            {
                foreach ($xmlLabel in $xmlTerm.Element("Labels").Elements("Label"))
                {
                    #We ignore the first True Label as this is the default label.
                    if ([bool]::Parse($xmlLabel.Attribute("IsDefaultForLanguage").Value) -ne $true)
                    {
                        #We also want to avoid adding duplicate labels
                        if (-Not $sTermLabels.Contains("|" + $xmlLabel.Attribute("Value").Value + "|"))
                        {
                            $mmsLabel = $mmsTerm.CreateLabel($xmlLabel.Attribute("Value").Value, [int]$xmlLabel.Attribute("Language").Value, [bool]::Parse($xmlLabel.Attribute("IsDefaultForLanguage").Value))
                            $bTermModified = $true

                            #We need to ensure we are not adding duplicate labels. This can happen with labels that have "&" in and the escaped version also "&amp;"
                            $sTermLabels += $xmlLabel.Attribute("Value").Value + "|"
                        }
                    }
                }
            }

            #Compare and check the XML and Term CustomProperties
            if ($xmlTerm.Element("CustomProperties") -ne $null)
            {
                foreach ($xmlCustProp in $xmlTerm.Element("CustomProperties").Elements("CustomProperty"))
                {
                    if (-Not $mmsTerm.CustomProperties.ContainsKey($xmlCustProp.Attribute("Key").Value))
                    {
                        $mmsTerm.SetCustomProperty($xmlCustProp.Attribute("Key").Value, $xmlCustProp.Attribute("Value").Value)
                        $bTermModified = $true
                    }
                }
            }

            #Compare and check the XML and Term LocalCustomProperties
            if ($xmlTerm.Element("LocalCustomProperties") -ne $null)
            {
                foreach ($xmlLocalCustProp in $xmlTerm.Element("LocalCustomProperties").Elements("LocalCustomProperty"))
                {
                    $mmsTerm.SetLocalCustomProperty($xmlLocalCustProp.Attribute("Key").Value, $xmlLocalCustProp.Attribute("Value").Value)
                    $bTermModified = $true
                }
            }

            #Check for a term move but not for the root nodes
            if ($mmsParentTerm -ne $null)
            {
                if ($mmsParentTerm.Id -ne $xmlTerm.Parent.Parent.Attribute("Id").Value)
                {
                    #We have a term move. We need to move to the XML definition
                    $mmsNewParentTerm = $mmsTermSet.GetTerm($xmlTerm.Parent.Parent.Attribute("Id").Value)
                    $spContext.Load($mmsNewParentTerm)
                    $mmsTerm = $mmsTerm.Move($mmsNewParentTerm)
                    $bTermModified = $true
                }
            }

            #We check here for any modifications and if we have one, execute the command
            if ($bTermModified)
            {
                try
                {
                    #$spContext.Load($mmsTerm);
                    $spContext.ExecuteQuery();
	                Write-Host "modified" -ForegroundColor White

                    #Update the XML Document "Processed" attribute in case we need to rerun the actions
                    $xmlTerm.Attribute("Processed").Value = $true
	            }
                catch
                {
                    Write-Host "Error occured while modifying a Term" $xmlTermName $_.Exception.Message -ForegroundColor Red
                
                    #Save the status
                    $xmlTerm.Document.Save($TermsFilePath)
                }
            }
            elseif (-not $xmlTermIsSourceTerm)
            {
                Write-Host "skipped - pin/re-used term" -ForegroundColor Red
            }
            else
            {
                Write-Host "not modified" -ForegroundColor Cyan
            
                #Update the XML Document "Processed" attribute in case we need to rerun the actions
                $xmlTerm.Attribute("Processed").Value = $true
            }
        }
    }
    else
    {
        #We have already processed this term
        Write-Host "previously processed" -ForegroundColor Cyan
    }
     
    if (!$errorOccurred)
    {
	    if ($xmlTerm.Element("Terms") -ne $null) 
        {
            foreach ($xmlChildTermNode in $xmlTerm.Element("Terms").Elements("Term")) 
            {
                Check-Term $xmlChildTermNode $false $mmsTermSet $mmsStore $lcid $spContext $global:glTermCurrentCount
            }
        }
    }
    else
    {
        #Save the status
        $xmlTerm.Document.Save($TermsFilePath)
    }
}


#####################################################################################################
#
# ConnectToSP 
# Makes the initial connection to SharePoint Admin portal
#
#####################################################################################################
function ConnectToSP($AdminUrl, $User, $securePassword){

    $spContext = New-Object Microsoft.SharePoint.Client.ClientContext($AdminUrl)

    if ($AdminUrl.Contains(".sharepoint.com")) #SharePoint Online
    {	
	    $spCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($User, $securePassword)
    }
    else #SharePoint On-premises
    {	
	    $networkCredentials = New-Object -TypeName System.Net.NetworkCredential
	    $networkCredentials.UserName = $User.Split('\')[1]
	    $networkCredentials.SecurePassword = $securePassword
	    $networkCredentials.Domain = $User.Split('\')[0]

	    [System.Net.CredentialCache]$spCredentials = New-Object -TypeName System.Net.CredentialCache
	    $AdminUri = [System.Uri]$AdminUrl
	    $spCredentials.Add($AdminUri, "NTLM", $networkCredentials)
    }

    #See if we can establish a connection
    $spContext.Credentials = $spCredentials
    $spContext.RequestTimeOut = 5000 * 60 * 10

    try
    {
        $spContext.ExecuteQuery()
        Write-Host "Established a connection to SharePoint on $AdminUrl" -ForegroundColor Green
    }
    catch
    {
        Write-Host "Not able to connect to SharePoint on $AdminUrl. Exception:$_.Exception.Message" -ForegroundColor red
        exit 1
    }
    return $spContext
}


#####################################################################################################
#
# This is where the code will run.
#
#####################################################################################################
#
#Inform the user how long the operation took.  This is the start date/time.
$dtStartTime = Get-Date

Write-Host "##################################################################" -ForegroundColor Cyan
Write-Host "Starting the MMS GROUP IMPORT at $dtStartTime" -ForegroundColor Cyan
Write-Host "##################################################################" -ForegroundColor Cyan

if ($AdminUser -eq "") {$AdminUser = Read-Host "What SharePoint Administrator user account, with Termstore Admin rights, do you want to use?"}
if ($AdminPassword.Length -eq 0) {[System.Security.SecureString]$AdminPassword = Read-Host -AsSecureString "What is the password for $AdminUser ?"}

if ($AdminUser.Contains("\"))
{
    $PathToSPClientdlls = "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI" #Default for SP on-premises       

    if ($AdminUrl -eq "") {$AdminUrl = Read-Host "What is the SharePoint Central Administration URL to IMPORT to?"} #If the user hasn't specified a URL, ask them for the tenant and fix-up
}
else
{
    $PathToSPClientdlls = "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI" #Default for SPO CSOM

    if ($AdminUrl -eq "")
    {
        $AdminUrl = Read-Host "Which TENANT do you want to IMPORT to?" #If the user hasn't specified a URL, ask them for the tenant and fix-up
        $AdminUrl = "https://$AdminUrl-admin.sharepoint.com"
    }
}

#-Update is a boolean
# if $false (Default) - the import will check each term in the XML file with the LIVE MMS values - This can take a long time, especially if the Export file used is up-to-date
# if $true - Then only changes will be implemented. It is assumed that all other terms are correct and complete
if ($Update -eq $null) {$Update = $false}

#Get the import file and re-request if the entered path is invalid
if ($TermsFilePath -eq "") {$TermsFilePath = Read-Host "Enter the full path to the XML import file" }
if (-not (Test-Path $TermsFilePath))
{
    do
    {
        $TermsFilePath = Read-Host "Cannot find the XML import file [$TermsFilePath]. Please RE-ENTER"
    } while (-not (Test-Path $TermsFilePath))
}

#Test for existence and add client dlls references and the PnP SharePoint PS Extensions
if (Test-Path "$PathToSPClientdlls\Microsoft.SharePoint.Client.dll")
{
    Add-Type -Path "$PathToSPClientdlls\Microsoft.SharePoint.Client.dll"
    Add-Type -Path "$PathToSPClientdlls\Microsoft.SharePoint.Client.Runtime.dll"
    Add-Type -Path "$PathToSPClientdlls\Microsoft.SharePoint.Client.Taxonomy.dll"
    Add-Type -Path "$PathToSPClientdlls\Microsoft.SharePoint.Client.DocumentManagement.dll"

    #Connect to SharePoint Online - If the user was requested for a password, it will already be a securestring
    if ($AdminPassword.ToString() -ne "System.Security.SecureString")
    {
        [System.Security.SecureString]$AdminPassword = ConvertTo-SecureString $AdminPassword -AsPlainText -Force
    }

    $spContext = ConnectToSP $AdminUrl $AdminUser $AdminPassword
    $mmsTermStore = Get-TermStoreInfo $spContext
    $xmlTermSetXML = Get-TermsToImport $TermsFilePath

    Check-Groups $spContext $mmsTermStore $xmlTermSetXML
}
else
{
    Write-Host "The path $PathToSPClientdlls does not point to the SharePoint CSOM libraries. Please correct and try again." -ForegroundColor Red
}

#Work out the elapsed time.
$dtEndTime = Get-Date
$dtDifference = New-TimeSpan -Start $dtStartTime -End $dtEndTime

Write-Host "##################################################################" -ForegroundColor Cyan
Write-Host "Completed..." -ForegroundColor Cyan
Write-Host "The Taxonomy import operation took $dtDifference" -ForegroundColor Cyan
Write-Host "##################################################################" -ForegroundColor Cyan