Function New-PSPivotTable {
   <#
      comment based help omitted here
   #>

   [cmdletbinding(DefaultParameterSetName="Property")]

   Param(
      [Parameter(Position=0,Mandatory=$True)]
      [object]$Inputobject,
      [Parameter()]
      [String]$yLabel,
      [Parameter(Mandatory=$True)]
      [String]$yProperty,
      [Parameter(Mandatory=$True)]
      [string]$xLabel,
      [Parameter(ParameterSetName="Property")]
      [string]$xProperty,
      [Parameter(ParameterSetName="Count")]
      [switch]$Count,
      [Parameter(ParameterSetName="Sum")]
      [string]$Sum,
      [Parameter(ParameterSetName="Sum")]
      [ValidateSet("None","KB","MB","GB","TB")]
      [string]$Format="None",
      [Parameter(ParameterSetName="Sum")]
      [ValidateScript({$_ -gt 0})]
      [int]$Round
   )

##### FUNCTION NEEDED FOR THE PIVOT TABLE #####
Begin {
     Write-Verbose "Starting $($myinvocation.mycommand)"
     $Activity="PS Pivot Table"
     $status="Creating new table"
     Write-Progress -Activity $Activity -Status $Status
     #initialize an array to hold results
     $result=@()
     #if no yLabel then use yProperty name
     if (-Not $yLabel) {
         $yLabel=$yProperty
     }
     Write-Verbose "Vertical axis label is $ylabel"
}
 Process {    
     Write-Progress -Activity $Activity -status "Pre-Processing"
     if ($Count -or $Sum) {
         #create an array of all unique property names so that if one isn’t 
         #found we can set a value of 0
         Write-Verbose "Creating a unique list based on $xLabel"
         <#
           Filter out blanks. Uniqueness is case sensitive so we first do a 
           quick filtering with Select-Object, then turn each of them to upper
           case and finally get unique uppercase items. 
         #>
         $unique=$inputobject | Where {$_.$xlabel} | 
          Select-Object -ExpandProperty $xLabel -unique | foreach {
            $_.ToUpper()} | Select-Object -unique
          
         Write-Verbose ($unique -join  ‘,’ | out-String).Trim()
       
     } 
     else {
      Write-Verbose "Processing $xLabel for $xProperty"    
     }
     
     Write-Verbose "Grouping objects on $yProperty"
     Write-Progress -Activity $Activity -status "Pre-Processing" -CurrentOperation "Grouping by $yProperty"
     $grouped=$Inputobject | Group -Property $yProperty
     $status="Analyzing data"  
     $i=0
     $groupcount=($grouped | measure).count
     foreach ($item in $grouped ) {
       Write-Verbose "Item $($item.name)"
       $i++
       #calculate what percentage is complete for Write-Progress
       $percent=($i/$groupcount)*100
       Write-Progress -Activity $Activity -Status $Status -CurrentOperation $($item.Name) -PercentComplete $percent
       $obj=new-object psobject -property @{$yLabel=$item.name}   
       #process each group
         #Calculate value depending on parameter set
         Switch ($pscmdlet.parametersetname) {
         
         "Property" {
                     <#
                       take each property name from the horizontal axis and make 
                       it a property name. Use the grouped property value as the 
                       new value
                     #>
                      $item.group | foreach {
                         $obj | Add-member Noteproperty -name "$($_.$xLabel)" -value $_.$xProperty
                       } #foreach
                     }
         "Count"  {
                     Write-Verbose "Calculating count based on $xLabel"
                      $labelGroup=$item.group | Group-Object -Property $xLabel 
                      #find non-matching labels and set count to 0
                      Write-Verbose "Finding 0 count entries"
                      #make each name upper case
                      $diff=$labelGroup | Select-Object -ExpandProperty Name -unique | 
                      Foreach { $_.ToUpper()} |Select-Object -unique
                      
                      #compare the master list of unique labels with what is in this group
                      Compare-Object -ReferenceObject $Unique -DifferenceObject $diff | 
                      Select-Object -ExpandProperty inputobject | foreach {
                         #add each item and set the value to 0
                         Write-Verbose "Setting $_ to 0"
                         $obj | Add-member Noteproperty -name $_ -value 0
                      }
                      
                      Write-Verbose "Counting entries"
                      $labelGroup | foreach {
                         $n=($_.name).ToUpper()
                         write-verbose $n
                         $obj | Add-member Noteproperty -name $n -value $_.Count -force
                     } #foreach
                  }
          "Sum"  {
                     Write-Verbose "Calculating sum based on $xLabel using $sum"
                     $labelGroup=$item.group | Group-Object -Property $xLabel 
                  
                      #find non-matching labels and set count to 0
                      Write-Verbose "Finding 0 count entries"
                      #make each name upper case
                      $diff=$labelGroup | Select-Object -ExpandProperty Name -unique | 
                      Foreach { $_.ToUpper()} |Select-Object -unique
                      
                      #compare the master list of unique labels with what is in this group
                      Compare-Object -ReferenceObject $Unique -DifferenceObject $diff | 
                      Select-Object -ExpandProperty inputobject | foreach {
                         #add each item and set the value to 0
                         Write-Verbose "Setting $_ sum to 0"
                         $obj | Add-member Noteproperty -name $_ -value 0
                      }
                      
                      Write-Verbose "Measuring entries"
                      $labelGroup | foreach {
                         $n=($_.name).ToUpper()
                         write-verbose "Measuring $n"
                         
                         $measure= $_.Group | Measure-Object -Property $Sum -sum
                         if ($Format -eq "None") {
                             $value=$measure.sum
                         }
                         else {
                             Write-Verbose "Formatting to $Format"
                              $value=$measure.sum/"1$Format"
                             }
                         if ($Round) {
                             Write-Verbose "Rounding to $Round places"
                             $Value=[math]::Round($value,$round)
                         }
                         $obj | Add-member Noteproperty -name $n -value $value -force
                     } #foreach
                 }        
         } #switch

        #add each object to the results array
       $result+=$obj
     } #foreach item
} #process
 End {
     Write-Verbose "Writing results to the pipeline"
     Return $result
     Write-Verbose "Ending $($myinvocation.mycommand)"
     Write-Progress -Completed -Activity $Activity -Status "Ending"
}
} #end function
##########