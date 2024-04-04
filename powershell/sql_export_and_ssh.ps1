#Potentially sensitive information has been scrubbed and replaced with placeholder data
#Intent of this script:  An in house integration between a SQL Server database in one isolated domain and a web server/database 
#                        on another domain for automation.  This half is executed in a SQL Agent job on the SQL Server.
#                        the other half for the web server/MySQL databaser portion is this file:

#NOSQLPS
Import-Module -Name SqlServer

# SQL Queries
# Workers
$WorkerSQL = "select p.PerKey as [extra_ID_1], --  !!! THIS IS IMPORTANT, IT IS OUR FOREIGN KEY BETWEEN tblPersonnel and EHSA..Worker
  p.EmployeeID as [WORKER_SOCIAL_SECURITY_NUMBER],
  wm.SSOID as [Institutional_ID],
  'SSO' as [Inst_ID_Type],
  IsNull(IsNull(wm.EmailAddress,wnm.NEmailAddress),'') as [WORKER_EMAIL_ADDRESS],
  p.PerFirstName as [WORKER_FIRST_NAME],
  IsNull(p.MiddleName,'') as [WORKER_MIDDLE_NAME],
  p.PerLastName as [WORKER_LAST_NAME],
  p.InformalName as [Worker_Preferred_Name],
  -- IN DEVELOPMENT: Build our worker structure in web app and determine how to identify those types with personnel database data.
  -- Worker Type as [WORKER_TYPE],
  case when p.EmploymentStatus != 8 then 'Active' else 'Inactive' end as [WORKER_STATUS],
  p.EmploymentStatus,
  es.EmploymentStatusName as extra_ID_2

from people_records as p
  left outer join work_records as wm
    on p.PerKey = wm.WrkPerKey
  left outer join other_work_records as wnm
    on p.PerKey = wnm.WrkNonPerKey
  left outer join employment_status as es
    on p.EmploymentStatus = es.EmpKey
  left outer join employee_access as a
    on p.PerKey = a.AccPerKey 

where p.EmploymentStatus not in ('8')				-- All employment status types minus inactive
  and a.AccessEnd >= getdate()-1					-- and their access hasn't ended since yesterday
  and wm.SSOID is not null							-- and they aren't missing an SSOID, which is key for user/training/dosimetry authentication
  -- and p.EmployeeID is not null						-- determine if non have an employee ID is also crucial, toggling this on 2/8/2024 removed 17 records

order by p.PerLastName"

# Users
$UserSQL = "select p.EmployeeID as [Social_Security_Number],
wm.SSOID as [USERID],
CONCAT(p.PerLastName,', ',p.PerFirstName) as [Name],
case when p.EmployeeID in ('person_id_1','person_id_2') then 'A' else 'EHS' end as Access,
-- IN DEVELOPMENT: Working with other departments to determine if and how we want to identifiy security groups from the personnel database.
--case when wm.[Group] = '5' and (IsSupervisor = 1 or wm.Title like '%Director%' or wm.Title like '%Manager%') then 'H&SSup, Training, Dosimetry'
   --  when wm.[Group] = '5' then 'H&S, Training, Dosimetry'
   --  when wm.[Group] = '6' and (IsSupervisor = 1 or wm.Title like '%Director%' or wm.Title like '%Manager%') then 'HPSup, Training, Dosimetry'
   --  when wm.[Group] = '6' then 'HP, Training, Dosimetry'
   --  when wm.[Group] = '21' then 'EHSAdmin, Training, Dosimetry'
   --  else 'Training, Dosimetry'
   --  end as Security_Groups
'Training, Dosimetry' as Security_Groups

from work_records as wm
left outer join employee_access..tblAccess as a
  on wm.WrkPerKey = a.AccPerKey
left outer join people_records as p
  on wm.WrkPerKey = p.perkey
left outer join divisions as d
  on wm.Division = d.DivKey
left outer join groups as g
  on wm.[Group] = g.GrpKey

where p.EmploymentStatus not in ('8')				-- All employment status types minus inactive
and a.AccessEnd >= getdate()-1					-- and their access hasn't ended since yesterday
and wm.SSOID is not null
and p.EmployeeID is not null -- and they aren't missing an SSOID, which is key for user/training/dosimetry authentication
--where wm.IsSupervisor = 1
--  and acc.AccessEnd >= getdate()-1
--where wm.[group] in ('5','6','21')
--  and acc.AccessEnd >= getdate()-1
--  and wm.SSOID is not null

order by p.PerLastName,
p.PerFirstName"

# Creation of DataSets, Executing SQL Statements and Populating Datasets
$SQLServer = "SQL Server"  
$SQLDBName = "Database"
$WorkerDataSet = New-Object System.Data.DataSet 
$UserDataSet = New-Object System.Data.DataSet
# Worker Dataset Population
$WorkerDataSet = Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDBName -Query $WorkerSQL -TrustServerCertificate -As DataSet  
for ($i = 0; $i -lt $WorkerDataSet.Tables.Count; $i++) {
    $WorkerCSV = $WorkerDataSet.Tables[$i] | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 
}
# User Dataset Population
$UserDataSet = Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDBName -Query $UserSQL -TrustServerCertificate -As DataSet  
for ($i = 0; $i -lt $UserDataSet.Tables.Count; $i++) {
    $UserCSV = $UserDataSet.Tables[$i] | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 
}

# Creating the files to be sent to the host(MURR-EHSA1)
# Worker file
$WorkerFileName = "WorkerExport"+$(Get-Date -Format "_yyyyHHmm_HHddss")+".csv"
$WorkerCSV | Out-File -FilePath C:\Temp\$WorkerFileName
$WorkerFile = Get-ChildItem -Path C:\Temp\ -Recurse -Include $WorkerFileName
$WorkerCSV = $WorkerFile.FullName
# User file
$UserFileName = "UserExport"+$(Get-Date -Format "_yyyyHHmm_HHddss")+".csv"
$UserCSV | Out-File -FilePath C:\Temp\$UserFileName
$UserFile = Get-ChildItem -Path C:\Temp\ -Recurse -Include $UserFileName
$UserCSV = $UserFile.FullName

# Open SSH connection and move file from client(MURR-SQL19) to host(MURR-EHSA1)
scp.exe -i ssh_key $WorkerCSV $UserCSV service_account@remote_server:C:\IntendedFileLocation 

# Removing local files for cleanliness
Remove-Item $WorkerCSV
Remove-Item $UserCSV
