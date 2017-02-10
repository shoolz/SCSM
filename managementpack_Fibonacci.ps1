#This script looks at all templates in in service manager, and pulls out defined fields that are set in the template
#Each field is stored into an array object for each class. the array is then dumped to excel for easy reading.
#new fields can be added using the xpath variables that are supposed by system center.


Remove-Variable * -ErrorAction SilentlyContinue
Import-Module SMLets

$server = [servername]


 if(!(test-path c:\temp\templates )) {New-Item c:\temp\templates -ItemType directory}


 function process-ma {
 Param($object)
   #found an MA, lets process it
            #loop throug MA nodes looking for AD node
            foreach($node in $object.ChildNodes)
            {
            
            $clean_node = $node.path.Replace("'", "")

             ########################### Title#########################
             #extracts MA title, here would be the place to desc or additional ma fields if desired
      If ('$Context/Property[Type=CustomSystem_WorkItem_Library!System.WorkItem]/Title$' -eq $clean_node)
      {
             
             $title = $node."#text"       
          $maobject.Title = "$Title"
                          
         
       }
                #if node is AD node
        if($ad_group_path -eq $clean_node)
            {
                #loop through AD node for displayname path
                foreach($ADNode in $node.childnodes)
                {
                $clean_node = $ADNode.path.replace("'","")
                    #once we find #displayname path , save text to variable
                    if($clean_node -eq $AD_displayname_path)
                    {
                    $assignedTo = $ADNode."#text"
                    $maobject.TemplateName = "$templatename"
                    $maobject.AssignedTo = "$assignedto"
                    $maobject  | Sort-Object -Property @{expression={$_.psobject.properties.count}} -Descending| Export-Csv c:\temp\templates\MAtemplates.csv -NoTypeInformation -append
                    }

                }


            }


            }

 }

#loops thorugh child xpaths and looks for any paths that match a manual activity, if found called process-ma
function checkfor-ma {
param($object)

  foreach($node in $object.childnodes)
  {
    $clean_node = $node.path.replace("'", "")
    if($ma_path -eq $clean_node){
      process-ma $node
    }
    
  }

}

#recursive function. this is the sauce that finds all nested MAs
function checkfor-nested {
param($object)

  foreach($node in $object.childnodes)
  {
    $clean_path = $node.path.replace("'", "")

 if($ma_path -eq $clean_path )
        {
          #process MA
          process-ma $node
        }

    if($pa_path -eq $clean_path){
     
      
      checkfor-nested $node
     
    }
        if($sa_path -eq $clean_path){
     
      
      checkfor-nested $node
     
    }
    
  }

}




$IRArray = @()
$CRArray = @()
$SRArray = @()
$PRArray = @()
$RRArray = @()
$MAArray = @()


$ma_path = '$Context/Path[Relationship=CustomSystem_WorkItem_Activity_Library!System.WorkItemContainsActivity TypeConstraint=CustomSystem_WorkItem_Activity_Library!System.WorkItem.Activity.ManualActivity]$'
$pa_path = '$Context/Path[Relationship=CustomSystem_WorkItem_Activity_Library!System.WorkItemContainsActivity TypeConstraint=CustomSystem_WorkItem_Activity_Library!System.WorkItem.Activity.ParallelActivity]$'
$sa_path = '$Context/Path[Relationship=CustomSystem_WorkItem_Activity_Library!System.WorkItemContainsActivity TypeConstraint=CustomSystem_WorkItem_Activity_Library!System.WorkItem.Activity.SequentialActivity]$'
$ad_group_path = '$Context/Path[Relationship=CustomSystem_WorkItem_Library!System.WorkItemAssignedToUser TypeConstraint=CustomMicrosoft_Windows_Library!Microsoft.AD.Group]$'
$AD_displayname_path = '$Context/Property[Type=CustomSystem_Library!System.Entity]/DisplayName$'
$templates = Get-SCSMObjectTemplate -ComputerName $server 
#$templates = Get-SCSMObjectTemplate -ComputerName $server | ? {$_.DisplayName -eq "Branch Move"} 

foreach($template in $templates)
{

  $CRObject = new-object System.Object
  $srobject = new-object System.Object
  $irObject = new-object System.Object
  $PRObject = new-object System.Object
  $RRObject = new-object System.Object
  $MAObject = new-object System.Object  

  [xml]$xml = $template.GetXML()
  
  #main work item properties
  $parent_properties = $xml.ObjectTemplate.Property
  #all child work item properties
  $top_objects = $xml.ObjectTemplate.Object
 
  $IR = $false
  $SR = $false
  $CR = $false
  $PR = $false
  $RR = $false
  
  $templateName = $template.DisplayName
  write-host "Trying template: $templatename"


  $crobject | Add-member -type NoteProperty -Name TemplateName -Value "$TemplateName"
  $Irobject | Add-member -type NoteProperty -Name TemplateName -Value "$TemplateName"
  $Srobject | Add-member -type NoteProperty -Name TemplateName -Value "$TemplateName"
  $Probject | Add-member -type NoteProperty -Name TemplateName -Value "$TemplateName"
  $Rrobject | Add-member -type NoteProperty -Name TemplateName -Value "$TemplateName"
  $MAobject | Add-member -type NoteProperty -Name TemplateName -Value "$TemplateName"
  
  
  $Probject | Add-member -type NoteProperty -name Title -Value "$Title"
  $Irobject | Add-member -type NoteProperty -name Title -Value "$Title"
  $Crobject | Add-member -type NoteProperty -name Title -Value "$Title"
  $Srobject | Add-member -type NoteProperty -name Title -Value "$Title"
  $Rrobject | Add-member -type NoteProperty -name Title -Value "$Title"
  $MAobject | Add-member -type NoteProperty -name Title -Value "$Title"
  
  
  $Probject | Add-member -type NoteProperty -name Description -Value "$desc"
  $Irobject | Add-member -type NoteProperty -name Description -Value "$desc"
  $Srobject | Add-member -type NoteProperty -name Description -Value "$desc"
  $Rrobject | Add-member -type NoteProperty -name Description -Value "$desc"
  $Crobject | Add-member -type NoteProperty -name Description -Value "$desc"
  
  $MAobject | Add-member -type NoteProperty -name AssignedTo -Value "$assigned"
  
  
  
    Foreach($prop in $parent_properties)
    {
      
      
      [string]$PathSTring = $prop.path
      $PathSTringS = $PathSTring.Replace("'", "")
            
     
      ########################### Title#########################
      If ('$Context/Property[Type=CustomSystem_WorkItem_Library!System.WorkItem]/Title$' -eq $PathSTringS)
      {
             
          
          [string]$output = $prop."#text"
          $split = $output.Split('=')
          $Title = $split[$split.Count-1].Replace("]", "").Replace("$", "").Replace("}", "").Replace("'", "")
          #Write-Host "Title: " $Name
          $Probject.Title = "$Title"
          $Irobject.Title = "$Title"
          $Crobject.Title = "$Title"
          $Srobject.Title = "$Title"
          $Rrobject.Title = "$Title"
         
                               
         
       }
      else{$title = ""}


      ########################## Description #######################
      If ('$Context/Property[Type=CustomSystem_WorkItem_Library!System.WorkItem]/Description$' -eq $PathSTringS)
      {
         
          [string]$output = $prop."#text"
          $split = $output.Split('=')
          $desc = $split[$split.Count-1].Replace("]", "").Replace("$", "").Replace("}", "").Replace("'", "")
          #Write-Host "Description: " $Name
          $Probject.Description = "$desc"
          $Irobject.Description = "$desc"
          $Srobject.Description = "$desc"
          $Rrobject.Description = "$desc"
          $Crobject.Description = "$desc"
        
         
         
      }
      else{$desc = ""}

      

      # Service Request

      If ('$Context/Property[Type=CustomSystem_WorkItem_ServiceRequest_Library!System.WorkItem.ServiceRequest]/Source$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
            
        if($output -match "mpelement")
        {
          #enum case 1
          $Name = $output.Replace("'", "").split("=")[1].split("!")[1].Replace("]", "").Replace("$", "").Replace("}", "")
          $Enum = Get-SCSMEnumeration -name "$Name" -ComputerName $server
          $source = $enum.displayname
        }
        else
        {
          #enum case 2
          $name = $output
          $Enum = Get-SCSMEnumeration -name "$Name" -ComputerName $server
          $source = $enum.displayname
            
        }
                               
        $Enum = Get-SCSMEnumeration -Name "$Name" -ComputerName $server
        #Write-Host "Source: " $Enum
        $srobject | Add-member -type NoteProperty -name Source -Value "$source"
        
        $SR = $true
            
      }
      If ('$Context/Property[Type=CustomSystem_WorkItem_ServiceRequest_Library!System.WorkItem.ServiceRequest]/Priority$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $Name = $output.Replace("'", "").split("!")[1].Replace("]", "").Replace("$", "").Replace("}", "")
        $Enum = Get-SCSMEnumeration -Name "$Name" -ComputerName $server
        #Write-Host "Priority: " $Enum
        $priority = $enum.displayname
        $srobject | Add-member -type NoteProperty -name Priority -Value "$priority"
        $SR = $true
      }
      If ('$Context/Property[Type=CustomSystem_WorkItem_ServiceRequest_Library!System.WorkItem.ServiceRequest]/Urgency$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $Name = $output.Replace("'", "").split("!")[1].Replace("]", "").Replace("$", "").Replace("}", "")
        $Enum = Get-SCSMEnumeration -Name "$Name" -ComputerName $server
        #Write-Host "Urgency: " $Enum
        $urgency = $enum.displayname
        $srobject | Add-member -type NoteProperty -name Urgency -Value "$urgency"
        $SR = $true
      }
      If ('$Context/Property[Type=CustomSystem_WorkItem_ServiceRequest_Library!System.WorkItem.ServiceRequest]/SupportGroup$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        
        #$split = $output.Split('=')
        #$Name = $split[$split.Count-1].Replace("]", "").Replace("$", "").Replace("}", "").Replace("'", "")
        $name = $output
        $Enum = Get-SCSMEnumeration -id "$Name" -ComputerName $server
        
        $SupportGroup = $enum.displayname
        $srobject | Add-member -type NoteProperty -name SupportGroup -Value "$supportgroup"
        $SR = $true
      }
      If ('$Context/Property[Type=CustomSystem_WorkItem_ServiceRequest_Library!System.WorkItem.ServiceRequest]/Area$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $split = $output.Split('!')
        $Name = $split[$split.count-1].Replace("]", "").Replace("$", "").Replace("}", "").Replace("'", "")
        $Name += "$"
        $Enum = Get-SCSMEnumeration -Name "$Name" -ComputerName $server
        #Write-Host "Area: " $Enum.displayname
        $area = $enum.displayname
        $srobject | Add-member -type NoteProperty -name Area -Value "$area"
        $SR = $true
      }
        
        
        
      # Incident
        
      If ('$Context/Property[Type=CustomSystem_WorkItem_Incident_Library!System.WorkItem.Incident]/Source$' -eq $PathSTringS)
      {
            
        [string]$output = $prop."#text"
        if($output -match "mpelement")
        {
          #enum case 1
          $Name = $output.Replace("'", "").split("=")[1].split("!")[1].Replace("]", "").Replace("$", "").Replace("}", "")
          $Enum = Get-SCSMEnumeration -name "$Name" -ComputerName $server
          $source = $enum.displayname
        }
        else
        {
          #enum case 2
          $name = $output
          $Enum = Get-SCSMEnumeration -name "$Name" -ComputerName $server
          $source = $enum.displayname
        }

        $irObject | Add-member -type NoteProperty -name Source -Value "$source"
        $ir = $true
      }
        
      If ('$Context/Property[Type=CustomSystem_WorkItem_Library!System.WorkItem.TroubleTicket]/Impact$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $Name = $output.Replace("'", "").split("!")[1].Replace("]", "").Replace("$", "").Replace("}", "")
        $Enum = Get-SCSMEnumeration -Name "$Name" -ComputerName $server
        #Write-Host "Impact: " $Enum
        $impact = $enum.displayname
        $irObject | Add-member -type NoteProperty -name Impact -Value "$impact"
        $ir = $true
      }
      If ('$Context/Property[Type=CustomSystem_WorkItem_Library!System.WorkItem.TroubleTicket]/Urgency$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $Name = $output.Replace("'", "").split("!")[1].Replace("]", "").Replace("$", "").Replace("}", "")
        $Enum = Get-SCSMEnumeration -Name "$Name" -ComputerName $server
        #Write-Host "Urgency: " $Enum
        $urgency = $enum.displayname
        $irObject | Add-member -type NoteProperty -name Urgency -Value "$Urgency"
        $ir = $true
      }
      If ('$Context/Property[Type=CustomSystem_WorkItem_Incident_Library!System.WorkItem.Incident]/TierQueue$' -eq $PathSTringS)
      {
        
        [string]$output = $prop."#text"
        $ID = $output
        if($id -ne 'a876be7d-7251-50b7-97eb-906a48e58733')
        {
          $Enum = Get-SCSMEnumeration -ID "$ID" -ComputerName $server
          $supportGroup = $enum.displayname
        }
        #Write-Host "Supportgroup; " $Enum.displayname
        else{$supportgroup = 'Blank'}
        $irObject | Add-member -type NoteProperty -name SupportGroup -Value "$supportgroup"
        $ir = $true
      }
      If ('$Context/Property[Type=CustomSystem_WorkItem_Incident_Library!System.WorkItem.Incident]/Classification$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $ID = $output
        $Enum = Get-SCSMEnumeration -ID "$ID" -ComputerName $server
        #Write-Host "Classification: " $Enum.displayname
        $classification = $enum.displayname
        $irObject | Add-member -type NoteProperty -name Classification -Value "$classification"
        $ir = $true
      }
               
      # Problem
        

      If ('$Context/Property[Type=CustomSystem_WorkItem_Problem_Library!System.WorkItem.Problem]/Source$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $ID = $output.Replace("'", "").split("=")[1].Replace("]", "").Replace("$", "").Replace("}", "")
        $Enum = Get-SCSMEnumeration -ID "$ID" -ComputerName $server
        #Write-Host "Source: " $Enum
        $probject | Add-member -type NoteProperty -name Source -Value $Enum
            
        $pr = $false
      }
        
      If ('$Context/Property[Type=CustomSystem_WorkItem_Problem_Library!System.WorkItem.Problem]/Classification$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $ID = $output.Split("=")[1].Replace("}", "")
        $Enum = Get-SCSMEnumeration -ID "$ID" -ComputerName $server
        #Write-Host "Classification: " $Enum.displayname
        $classification = $enum.displayname
        $probject | Add-member -type NoteProperty -name Classification -Value "$classification"
        $pr = $false
      }

      # Change
     

      If ('$Context/Property[Type=CustomSystem_WorkItem_ChangeRequest_Library!System.WorkItem.ChangeRequest]/Reason$' -eq $PathSTringS)
      {
      
      
        [string]$output = $prop."#text"
        $split = $output.Split('=')
        $Name = $split[$split.Count-1].Replace("]", "").Replace("$", "").Replace("}", "").Replace("'", "")
        #Write-Host "Reason: " $Name
        $CRobject | Add-member -type NoteProperty -name Reason -Value "$Name"

        $Cr = $true
      }
      If ('$Context/Property[Type=CustomSystem_WorkItem_ChangeRequest_Library!System.WorkItem.ChangeRequest]/Priority$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $split = $output.Split('=')
        $Name = $split[$split.Count-1].Replace("]", "").Replace("$", "").Replace("}", "").Replace("'", "")
        $Enum = Get-SCSMEnumeration -ID "$Name" -ComputerName $server
        #Write-Host "Priority: " $Enum
        $priority = $enum.displayname
        $CRobject | Add-member -type NoteProperty -name Priority -Value "$priority"
        $Cr = $true
      }
      If ('$Context/Property[Type=CustomSystem_WorkItem_ChangeRequest_Library!System.WorkItem.ChangeRequest]/Impact$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $split = $output.Split('=')
        $Name = $split[$split.Count-1].Replace("]", "").Replace("$", "").Replace("}", "").Replace("'", "")
        $Enum = Get-SCSMEnumeration -ID "$Name" -ComputerName $server
        #Write-Host "Impact: " $Enum
        $impact = $enum.displayname
        $CRobject | Add-member -type NoteProperty -name Impact -Value "$impact"
        $Cr = $true
      }
      If ('$Context/Property[Type=CustomSystem_WorkItem_ChangeRequest_Library!System.WorkItem.ChangeRequest]/Risk$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $split = $output.Split('=')
        $Name = $split[$split.Count-1].Replace("]", "").Replace("$", "").Replace("}", "").Replace("'", "")
        $Enum = Get-SCSMEnumeration -ID "$Name" -ComputerName $server
        #Write-Host "Risk; " $Enum.displayname
        $risk = $enum.displayname
        $CRobject | Add-member -type NoteProperty -name Risk -Value "$risk"
        $Cr = $true
      }
      If ('$Context/Property[Type=CustomSystem_WorkItem_ChangeRequest_Library!System.WorkItem.ChangeRequest]/Area$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $split = $output.Split('=')
        $Name = $split[$split.Count-1].Replace("]", "").Replace("$", "").Replace("}", "").Replace("'", "")
        $Enum = Get-SCSMEnumeration -ID "$Name" -ComputerName $server
        $area = $enum.displayname
        #Write-Host "Area: " $Enum.displayname
        $CRobject | Add-member -type NoteProperty -name Area -Value "$area"
        $Cr = $true
      }
     
      If ('$Context/Property[Type=CustomSystem_WorkItem_ChangeRequest_Library!System.WorkItem.ChangeRequest]/ImplementationPlan$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $split = $output.Split('=')
        $Name = $split[$split.Count-1].Replace("]", "").Replace("$", "").Replace("}", "").Replace("'", "")
        #Write-Host "Reason: " $Name
        $CRobject | Add-member -type NoteProperty -name ImplementationPlan -Value "$Name"
        $Cr = $true
      }
      If ('$Context/Property[Type=CustomSystem_WorkItem_ChangeRequest_Library!System.WorkItem.ChangeRequest]/RiskAssessmentPlan$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $split = $output.Split('=')
        $Name = $split[$split.Count-1].Replace("]", "").Replace("$", "").Replace("}", "").Replace("'", "")
        #Write-Host "Reason: " $Name
        $CRobject | Add-member -type NoteProperty -name RiskAssessmentPlan -Value "$Name"
        $Cr = $true
      }
      If ('$Context/Property[Type=CustomSystem_WorkItem_ChangeRequest_Library!System.WorkItem.ChangeRequest]/BackoutPlan$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $split = $output.Split('=')
        $Name = $split[$split.Count-1].Replace("]", "").Replace("$", "").Replace("}", "").Replace("'", "")
        #Write-Host "Reason: " $Name
        $CRobject | Add-member -type NoteProperty -name BackOutPlan -Value "$Name"
        $Cr = $true
      }
      If ('$Context/Property[Type=CustomSystem_WorkItem_ChangeRequest_Library!System.WorkItem.ChangeRequest]/TestPlan$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $split = $output.Split('=')
        $Name = $split[$split.Count-1].Replace("]", "").Replace("$", "").Replace("}", "").Replace("'", "")
        #Write-Host "Reason: " $Name
        $CRobject | Add-member -type NoteProperty -name TestPlan -Value "$Name"
        $Cr = $true
      }    
     
      If ('$Context/Property[Type=CustomSystem_WorkItem_ChangeRequest_Library!System.WorkItem.ChangeRequest]/Category$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $split = $output.Split('=')
        $Name = $split[$split.Count-1].Replace("]", "").Replace("$", "").Replace("}", "").Replace("'", "")
        $Enum = Get-SCSMEnumeration -ID "$Name" -ComputerName $server
        $category = $enum.displayname
        #Write-Host "Area: " $Enum.displayname
        $CRobject | Add-member -type NoteProperty -name Category -Value "$category"
        $Cr = $true
      }  
      If ('$Context/Property[Type=CustomStifel_ChangeFormCustomizations!ClassExtension_2e27121d_21df_4261_9fbf_adc727bf66c8]/DataCenterAccess$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $split = $output.Split('=')
        $Name = $split[$split.Count-1].Replace("]", "").Replace("$", "").Replace("}", "").Replace("'", "")
        #Write-Host "Reason: " $Name
        $CRobject | Add-member -type NoteProperty -name DataCenterAccess -Value "$Name"
        $Cr = $true
      } 
      If ('$Context/Property[Type=CustomStifel_ChangeFormCustomizations!ClassExtension_2e27121d_21df_4261_9fbf_adc727bf66c8]/CR_Ext05_Long$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $split = $output.Split('=')
        $Name = $split[$split.Count-1].Replace("]", "").Replace("$", "").Replace("}", "").Replace("'", "")
        #Write-Host "Reason: " $Name
        $CRobject | Add-member -type NoteProperty -name CommunicationPlan -Value "$Name"
        $Cr = $true
      } 
      If ('$Context/Property[Type=CustomStifel_ChangeFormCustomizations!ClassExtension_2e27121d_21df_4261_9fbf_adc727bf66c8]/CR_Ext16$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $split = $output.Split('=')
        $Name = $split[$split.Count-1].Replace("]", "").Replace("$", "").Replace("}", "").Replace("'", "")
        #Write-Host "Reason: " $Name
        $CRobject | Add-member -type NoteProperty -name DEV -Value "$Name"
        $Cr = $true
      }
      If ('$Context/Property[Type=CustomStifel_ChangeFormCustomizations!ClassExtension_2e27121d_21df_4261_9fbf_adc727bf66c8]/CR_Ext17$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $split = $output.Split('=')
        $Name = $split[$split.Count-1].Replace("]", "").Replace("$", "").Replace("}", "").Replace("'", "")
        #Write-Host "Reason: " $Name
        $CRobject | Add-member -type NoteProperty -name QA -Value "$Name"
        $Cr = $true
      } 
      If ('$Context/Property[Type=CustomStifel_ChangeFormCustomizations!ClassExtension_2e27121d_21df_4261_9fbf_adc727bf66c8]/CR_Ext18$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $split = $output.Split('=')
        $Name = $split[$split.Count-1].Replace("]", "").Replace("$", "").Replace("}", "").Replace("'", "")
        #Write-Host "Reason: " $Name
        $CRobject | Add-member -type NoteProperty -name UAT -Value "$Name"
        $Cr = $true
      } 
      If ('$Context/Property[Type=CustomStifel_ChangeFormCustomizations!ClassExtension_2e27121d_21df_4261_9fbf_adc727bf66c8]/CR_Ext19$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $split = $output.Split('=')
        $Name = $split[$split.Count-1].Replace("]", "").Replace("$", "").Replace("}", "").Replace("'", "")
        #Write-Host "Reason: " $Name
        $CRobject | Add-member -type NoteProperty -name Prod -Value "$Name"
        $Cr = $true
      } 
      If ('$Context/Property[Type=CustomStifel_ChangeFormCustomizations!ClassExtension_2e27121d_21df_4261_9fbf_adc727bf66c8]/CR_Ext20$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $split = $output.Split('=')
        $Name = $split[$split.Count-1].Replace("]", "").Replace("$", "").Replace("}", "").Replace("'", "")
        #Write-Host "Reason: " $Name
        $CRobject | Add-member -type NoteProperty -name DR -Value "$Name"
        $Cr = $true
      } 
      If ('$Context/Property[Type=CustomStifel_ChangeFormCustomizations!ClassExtension_2e27121d_21df_4261_9fbf_adc727bf66c8]/CR_MaintenanceWindow$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $split = $output.Split('=')
        $Name = $split[$split.Count-1].Replace("]", "").Replace("$", "").Replace("}", "").Replace("'", "")
        #Write-Host "Reason: " $Name
        $CRobject | Add-member -type NoteProperty -name MaintenanceWindow -Value "$Name"
        $Cr = $true
      } 
      If ('$Context/Property[Type=CustomStifel_ChangeFormCustomizations!ClassExtension_2e27121d_21df_4261_9fbf_adc727bf66c8]/Reminder$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $split = $output.Split('=')
        $Name = $split[$split.Count-1].Replace("]", "").Replace("$", "").Replace("}", "").Replace("'", "")
        #Write-Host "Reason: " $Name
        $CRobject | Add-member -type NoteProperty -name Reminder -Value "$Name"
        $Cr = $true
      } 
       

      # Release

      If ('$Context/Property[Type=CustomSystem_WorkItem_ReleaseRecord_Library!System.WorkItem.ReleaseRecord]/Notes$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $split = $output.Split('=')
        $Name = $split[$split.Count-1].Replace("]", "").Replace("$", "").Replace("}", "").Replace("'", "")
        #Write-Host "Notes: " $Name
        $RRobject | Add-member -type NoteProperty -name Notes -Value $Name
            
        $rr = $true
      }
      If ('$Context/Property[Type=CustomSystem_WorkItem_ReleaseRecord_Library!System.WorkItem.ReleaseRecord]/Type$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $split = $output.Split('!')
        $Name = $split[$split.Count-1].Replace("]", "").Replace("$", "").Replace("}", "").Replace("'", "")
        $Enum = Get-SCSMEnumeration -Name "$Name" -ComputerName $server
        #Write-Host "Type: " $Enum
        $RRobject | Add-member -type NoteProperty -name Type -Value $Enum
        $rr = $true
      }
      If ('$Context/Property[Type=CustomSystem_WorkItem_ReleaseRecord_Library!System.WorkItem.ReleaseRecord]/Category$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $split = $output.Split('!')
        $Name = $split[$split.Count-1].Replace("]", "").Replace("$", "").Replace("}", "").Replace("'", "")
        $Enum = Get-SCSMEnumeration -Name "$Name" -ComputerName $server
        #Write-Host "Category: " $Enum
        $RRobject | Add-member -type NoteProperty -name Category -Value $Enum
        $rr = $true
      }
      If ('$Context/Property[Type=CustomSystem_WorkItem_ReleaseRecord_Library!System.WorkItem.ReleaseRecord]/Impact$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $split = $output.Split('!')
        $Name = $split[$split.Count-1].Replace("]", "").Replace("$", "").Replace("}", "").Replace("'", "")
        $Enum = Get-SCSMEnumeration -Name "$Name" -ComputerName $server
        #Write-Host "Impact: " $Enum
        $RRobject | Add-member -type NoteProperty -name Impact -Value $Enum
        $rr = $true
      }
      If ('$Context/Property[Type=CustomSystem_WorkItem_ReleaseRecord_Library!System.WorkItem.ReleaseRecord]/Risk$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $split = $output.Split('!')
        $Name = $split[$split.Count-1].Replace("]", "").Replace("$", "").Replace("}", "").Replace("'", "")
        $Enum = Get-SCSMEnumeration -Name "$Name" -ComputerName $server
        #Write-Host "Risk: " $Enum
        $RRobject | Add-member -type NoteProperty -name Risk -Value $Enum
        $rr = $true
      }
      If ('$Context/Property[Type=CustomSystem_WorkItem_ReleaseRecord_Library!System.WorkItem.ReleaseRecord]/Priority$' -eq $PathSTringS)
      {
        [string]$output = $prop."#text"
        $split = $output.Split('!')
        $Name = $split[$split.Count-1].Replace("]", "").Replace("$", "").Replace("}", "").Replace("'", "")
        $Enum = Get-SCSMEnumeration -Name "$Name" -ComputerName $server
        #Write-Host "Priority; " $Enum.displayname
        $RRobject | Add-member -type NoteProperty -name Priority -Value $Enum
        $rr = $true
      }
        
          
      #end main node loop 
    }  
     


     foreach($object in $top_objects)
{

     $clean_path = $object.Path.Replace("'", "")
     
    #if any top paths are ma's, process them
          if($ma_path -eq $clean_path )
        {
          #process MA
          process-ma $object
        }

           if($pa_path -eq $clean_path )
        {
         checkfor-ma $object
         checkfor-nested $object
          
        }

     

}
        
    
    

  if($pr)
  {


    $prArray += $PRObject
  }
  if($Cr)
  {
    $crarray += $CRObject
  }
  if($SR)
  {
    $SRArray += $srobject
  }
  if($RR)
  {
    $rrArray += $RRobject
  }
  if($ir)
  {
    $irArray += $irObject
  }

      #end template loop
      }


$CRarray | Sort-Object -Property @{expression={$_.psobject.properties.count}} -Descending| Export-Csv c:\temp\templates\CRtemplates.csv -NoTypeInformation
$SRarray | Sort-Object -Property @{expression={$_.psobject.properties.count}} -Descending| Export-Csv c:\temp\templates\sRtemplates.csv -NoTypeInformation
$IRarray | Sort-Object -Property @{expression={$_.psobject.properties.count}} -Descending| Export-Csv c:\temp\templates\iRtemplates.csv -NoTypeInformation
$PRarray | Sort-Object -Property @{expression={$_.psobject.properties.count}} -Descending| Export-Csv c:\temp\templates\pRtemplates.csv -NoTypeInformation
$RRarray | Sort-Object -Property @{expression={$_.psobject.properties.count}} -Descending| Export-Csv c:\temp\templates\rRtemplates.csv -NoTypeInformation