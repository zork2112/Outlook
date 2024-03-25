Attribute VB_Name = "MoveMessages"
Public lastError As Date
Private Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "User32" () As Long
Private Declare PtrSafe Function CloseClipboard Lib "User32" () As Long

Sub MoveToABGConversion()
    
    MoveMessage ("\Action\ABG\Relius Conversion")
    
End Sub

Sub MoveToACH()

  MoveMessage ("\Action\ACH")
  
End Sub
Sub MoveToAcumatica()

  MoveMessage ("\Action\Billing\Acumatica")
  
End Sub
Sub MoveToAdvisor()

  MoveMessage ("\Action\Advisor")
  
End Sub

Sub MoveToAdvisorBilling()

  MoveMessage ("\Action\Billing\Advisor")
  
End Sub

Sub MoveToBrightscope()

  MoveMessage ("\Action\BrightScope")
  
End Sub
Sub MoveToApproval()

  MoveMessage ("\Action\Approval")
  
End Sub
Sub MoveToAzure()

  MoveMessage ("\Action\Billing\Azure")
  
End Sub
Sub MoveToBenefits()

  MoveMessage ("\Personal\Benefits")
  
End Sub
Sub MoveToBilling()
    
    MoveMessage ("\Action\Billing")
    
End Sub

Sub MoveToBillingEmailAlerts()

  MoveMessage ("\Action\Billing\Email Alerts")
  
End Sub
Sub MoveToBIS()
    
    MoveMessage ("\Action\BIS")
    
End Sub
Sub MoveToCashbook()

  MoveMessage ("\Action\Cashbook/Checkbook")
  
End Sub
Sub MoveToCompliance()

  MoveMessage ("\Action\Compliance Testing")
  
End Sub
Sub MoveToCompanyEmailBlasts()
    
    MoveMessage ("\Action\Company Email Blasts")
    
End Sub
Sub MoveToCompanyPolicy()

  MoveMessage ("\Personal\Company Policies")
  
End Sub
Sub MoveToConsultants()
    
    MoveMessage ("\Action\Consultants")
    
End Sub
Sub MoveToCopado()

  MoveMessage ("\Action\Salesforce\Copado")
  
End Sub

Sub MoveToCopadoPromotion()

  MoveMessage ("\Action\Salesforce\Copado\Copado Promotion")
  
End Sub

Sub MoveToCRM()

  MoveMessage ("\Action\Pentegra Modernization Program\CRM")
  
End Sub

Sub MoveToCulture()
    
    MoveMessage ("\Action\Culture")
    
End Sub
Sub MoveToDataCleanup()

  MoveMessage ("\Action\Pentegra Modernization Program\Data Cleanup")
  
End Sub
Sub MoveToDataGovernance()

  MoveMessage ("\Action\Data Governance")
  
End Sub
Sub MoveToDataRequest()

  MoveMessage ("\Action\Data Request")
  
End Sub

Sub MoveToDataScience()

  MoveMessage ("\Personal\Training\Data Science")
  
End Sub

Sub MoveToDBA()

  MoveMessage ("\Action\DBA & SQL Issues")
  
End Sub
Sub MoveToDBWeb()

  MoveMessage ("\Action\DB Web - BlueRush")
  
End Sub
Sub MoveToDelete90()

    MoveMessage ("\Delete in 90 days")

End Sub
Sub MoveToDenodo()
    
    MoveMessage ("\Action\Denodo")
    
End Sub
Sub MoveToDocumentation()
    
    MoveMessage ("\Action\Documentation")
    
End Sub

Sub MoveToDR()
    
    MoveMessage ("\Action\Disaster Recovery")
    
End Sub
Sub MoveToEcosystem()
    
    MoveMessage ("\Action\Ecosystem")
    
End Sub
Sub MoveToEmailDbMail()
    
    MoveMessage ("\Action\Email DB Mail")
    
End Sub
Sub MoveToEMPI()

  MoveMessage ("\Action\Pentegra Modernization Program\Data Cleanup\EMPI")
  
End Sub
Sub MoveToEmpower()

  MoveMessage ("\Action\Empower")
  
End Sub
Sub MoveToFundPerformance()
    
    MoveMessage ("\Action\Fund Performance")
    
End Sub
Sub MoveToInbox()

  MoveMessage ("\Inbox")
  
End Sub
Sub MoveToInboxReview()

  MoveMessage ("\Inbox\Review")
  
End Sub
Sub MoveToInfrastructure()

  MoveMessage ("\Action\Infrastructure")
  
End Sub
Sub MoveToJobs()
    
    MoveMessage ("\Action\Job Results")
    
End Sub
Sub MoveToKeep()

    MoveMessage ("\Keep Forever")

End Sub
Sub MoveToMetLife()

  MoveMessage ("\Action\Met Life Conversion of Stable")
  
End Sub
Sub MoveToNewkirk()

  MoveMessage ("\Action\Newkirk")
  
End Sub
Sub MoveToPencal()

  MoveMessage ("\Action\Pencal")
  
End Sub
Sub MoveToPensionPro()

  MoveMessage ("\Action\PensionPro")
  
End Sub
Sub MoveToPersonal()

  MoveMessage ("\Personal")
  
End Sub

Sub MoveToPlanGenerator()
    
    MoveMessage ("\Action\Plan Generator")
    
End Sub

Sub MoveToPlanning()
    
    MoveMessage ("\Action\Planning")
    
End Sub
Sub MoveToProgramming()
    
    MoveMessage ("\Action\Programming")
    
End Sub
Sub MoveToPSI19Q4Updates()
    
    MoveMessage ("\Invoice\PSI 2019 Q4\PSI 19 Q4 Updates")
    
End Sub
Sub MoveToRelationshipVelocity()

  MoveMessage ("\Action\Pentegra Modernization Program")
  
End Sub
Sub MoveToReporting()

  MoveMessage ("\Action\Reporting")
  
End Sub
Sub MoveToReview()

  MoveMessage ("\Review")
  
End Sub
Sub MoveToRevision()

    MoveMessage ("\Invoice\PSI 2019 Q1\_Revisions")

End Sub
Sub MoveToRVBilling()

  MoveMessage ("\Action\Pentegra Modernization Program\Billing")
  
End Sub
Sub MoveToRVWeb()

  MoveMessage ("\Action\Pentegra Modernization Program\Web")
  
End Sub
Sub MoveToSalesforce()

  MoveMessage ("\Action\Salesforce")
  
End Sub
Sub MoveToSalesforceIntegration()

  MoveMessage ("\Action\Salesforce\Pipeline - Integration")
  
End Sub

Sub MoveToSchwab()

  MoveMessage ("\Action\Schwab")
  
End Sub
Sub MoveToSharePoint()
    
    MoveMessage ("\Action\SharePoint")
    
End Sub
Sub MoveToSnooze()

  MoveMessage ("\Inbox\Snooze")
  
End Sub
Sub MoveToSQLBeacon()

  MoveMessage ("\Action\SQLBeacon")
  
End Sub
Sub MoveToSRTASP()

  MoveMessage ("\Action\SRT ASP")
  
End Sub

Sub MoveToSRTUpgrade()
    
    MoveMessage ("\Action\Schwab\SRT Upgrade")
    
End Sub

Sub MoveToStorage()

    MoveMessage ("\Storage")
    
End Sub
Sub MoveToStraightPath()

  MoveMessage ("\Action\StraightPath")
  
End Sub
Sub MoveToTeams()
    
    MoveMessage ("\Inbox\_Triage\Teams")
    
End Sub

Sub MoveToTFSActivity()

  MoveMessage ("\Action\TFS Activity")
  
End Sub

Sub MoveToTpaTaskForce()

    MoveMessage ("\Action\Billing\Tpa Task Force")
    
End Sub
Sub MoveToTrading()
    
    MoveMessage ("\Action\Trading")
    
End Sub
Sub MoveToTraining()

  MoveMessage ("\Action\Training")
  
End Sub
Sub MoveToTriage()
    
    MoveMessage ("\Inbox\_Triage")
    
End Sub
Sub MoveToWeb()
    
    MoveMessage ("\Action\Web")
    
End Sub
Sub MoveToWorkfront()

  MoveMessage ("\Action\Pentegra Modernization Program\Workfront")
  
End Sub
Sub MoveToZennify()

  MoveMessage ("\Action\Zennify")
  
End Sub
Sub MoveToZennifyNewOrg()

  MoveMessage ("\Action\Salesforce\Zennify New Org")
  
End Sub



Sub MoveMessage(ByVal oFolderName As String)
  
 
  ' DoneFolder is of type Outlook.MAPIFolder
  
    oFolderPath = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Parent.FolderPath + oFolderName
  
  ' 20221026 had to add this due to the disconnect between my account and actual email :(
  ' 20221026 later that day Outlook decided to show rob.scholl - Oi!
  ' oFolderPath = Replace(oFolderPath, "rob.scholl@pentegra", "robert.scholl@pentegra")
  
    Set DoneFolder = GetFolder(oFolderPath)
  
   If oFolderName = "Action" Or oFolderName = "Waiting" Then InfoToClipboard
     
    For Each msg In ActiveExplorer.Selection
     
     msg.UnRead = False
     
     If oFolderName <> "Active" Then
        
            msg.FlagIcon = olNoFlagIcon
            msg.FlagStatus = olFlagComplete
            
        End If
        
     
     ' Msg.FlagRequest = "Completed"
     ' Msg.ReminderSet = False
     msg.Move DoneFolder
     Next msg
     
    

End Sub
