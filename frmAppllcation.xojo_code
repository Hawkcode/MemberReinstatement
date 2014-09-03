#tag WebPage
Begin WebPage frmAppllcation
   Compatibility   =   ""
   Cursor          =   0
   Enabled         =   True
   Height          =   501
   HelpTag         =   ""
   HorizontalCenter=   0
   ImplicitInstance=   True
   Index           =   0
   IsImplicitInstance=   False
   Left            =   0
   LockBottom      =   False
   LockHorizontal  =   False
   LockLeft        =   False
   LockRight       =   False
   LockTop         =   False
   LockVertical    =   False
   MinHeight       =   450
   MinWidth        =   600
   Style           =   "997821280"
   TabOrder        =   0
   Title           =   "Membership Application"
   Top             =   0
   VerticalCenter  =   0
   Visible         =   True
   Width           =   950
   ZIndex          =   1
   _DeclareLineRendered=   False
   _HorizontalPercent=   0.0
   _ImplicitInstance=   False
   _IsEmbedded     =   False
   _Locked         =   False
   _NeedsRendering =   True
   _OfficialControl=   False
   _OpenEventFired =   False
   _ShownEventFired=   False
   _VerticalPercent=   0.0
   Begin WebProgressBar prgProgress
      Cursor          =   0
      Enabled         =   True
      Height          =   20
      HelpTag         =   ""
      HorizontalCenter=   0
      Indeterminate   =   True
      Index           =   -2147483648
      Left            =   831
      LockBottom      =   False
      LockedInPosition=   False
      LockHorizontal  =   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      LockVertical    =   False
      Maximum         =   0
      Scope           =   0
      Style           =   "-1"
      TabOrder        =   -1
      Top             =   473
      Value           =   100
      VerticalCenter  =   0
      Visible         =   False
      Width           =   83
      ZIndex          =   1
      _DeclareLineRendered=   False
      _HorizontalPercent=   0.0
      _IsEmbedded     =   False
      _Locked         =   False
      _NeedsRendering =   True
      _OfficialControl=   False
      _OpenEventFired =   False
      _VerticalPercent=   0.0
   End
   Begin WebLabel Label4
      Cursor          =   1
      Enabled         =   True
      HasFocusRing    =   True
      Height          =   22
      HelpTag         =   ""
      HorizontalCenter=   0
      Index           =   -2147483648
      Left            =   171
      LockBottom      =   False
      LockedInPosition=   False
      LockHorizontal  =   True
      LockLeft        =   False
      LockRight       =   False
      LockTop         =   False
      LockVertical    =   True
      Multiline       =   False
      Scope           =   0
      Style           =   "816938816"
      TabOrder        =   0
      Text            =   "American Society of Plumbing Engineers Reinstatement Application"
      Top             =   14
      VerticalCenter  =   0
      Visible         =   True
      Width           =   607
      ZIndex          =   1
      _DeclareLineRendered=   False
      _HorizontalPercent=   0.0
      _IsEmbedded     =   False
      _Locked         =   False
      _NeedsRendering =   True
      _OfficialControl=   False
      _OpenEventFired =   False
      _VerticalPercent=   0.0
   End
   Begin WebButton btnPrevious
      Caption         =   "<Previous"
      Cursor          =   0
      Enabled         =   True
      Height          =   22
      HelpTag         =   ""
      HorizontalCenter=   0
      Index           =   -2147483648
      Left            =   614
      LockBottom      =   False
      LockedInPosition=   False
      LockHorizontal  =   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      LockVertical    =   False
      Scope           =   0
      Style           =   "0"
      TabOrder        =   11
      Top             =   473
      VerticalCenter  =   0
      Visible         =   False
      Width           =   100
      ZIndex          =   1
      _DeclareLineRendered=   False
      _HorizontalPercent=   0.0
      _IsEmbedded     =   False
      _Locked         =   False
      _NeedsRendering =   True
      _OfficialControl=   False
      _OpenEventFired =   False
      _VerticalPercent=   0.0
   End
   Begin WebTimer Timer1
      Cursor          =   0
      Enabled         =   True
      Height          =   32
      HelpTag         =   ""
      HorizontalCenter=   0
      Index           =   -2147483648
      Left            =   20
      LockBottom      =   False
      LockedInPosition=   False
      LockHorizontal  =   False
      LockLeft        =   False
      LockRight       =   False
      LockTop         =   False
      LockVertical    =   False
      Mode            =   0
      Period          =   1000
      Scope           =   0
      Style           =   "-1"
      TabOrder        =   -1
      TabPanelIndex   =   0
      Top             =   4
      VerticalCenter  =   0
      Visible         =   True
      Width           =   32
      ZIndex          =   1
      _DeclareLineRendered=   False
      _HorizontalPercent=   0.0
      _IsEmbedded     =   False
      _Locked         =   False
      _NeedsRendering =   True
      _OfficialControl=   False
      _OpenEventFired =   False
      _VerticalPercent=   0.0
   End
   Begin SMTPServer SMTPServer1
      CertificateFile =   
      CertificatePassword=   ""
      CertificateRejectionFile=   
      ConnectionType  =   2
      Height          =   32
      Index           =   -2147483648
      Left            =   20
      LockedInPosition=   False
      Scope           =   0
      Secure          =   False
      Style           =   "-1"
      TabPanelIndex   =   0
      Top             =   4
      Width           =   32
   End
   Begin dlgCorrectPerson dlgCorrectPersonYN
      Cursor          =   0
      Enabled         =   True
      Height          =   76
      HelpTag         =   ""
      HorizontalCenter=   0
      Index           =   -2147483648
      Left            =   20
      LockBottom      =   False
      LockedInPosition=   False
      LockHorizontal  =   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      LockVertical    =   False
      mbYes           =   False
      MinHeight       =   0
      MinWidth        =   0
      Resizable       =   True
      Scope           =   0
      Style           =   "-1"
      TabOrder        =   -1
      TabPanelIndex   =   0
      Title           =   "Untitled"
      Top             =   20
      Type            =   1
      VerticalCenter  =   0
      Visible         =   True
      Width           =   211
      ZIndex          =   1
      _DeclareLineRendered=   False
      _HorizontalPercent=   0.0
      _IsEmbedded     =   False
      _Locked         =   False
      _NeedsRendering =   True
      _OfficialControl=   False
      _OpenEventFired =   False
      _ShownEventFired=   False
      _VerticalPercent=   0.0
   End
   Begin WebTimer TimerQuit
      Cursor          =   0
      Enabled         =   True
      Height          =   32
      HelpTag         =   ""
      HorizontalCenter=   0
      Index           =   -2147483648
      Left            =   60
      LockBottom      =   False
      LockedInPosition=   False
      LockHorizontal  =   False
      LockLeft        =   False
      LockRight       =   False
      LockTop         =   False
      LockVertical    =   False
      Mode            =   0
      Period          =   2000
      Scope           =   0
      Style           =   "-1"
      TabOrder        =   -1
      TabPanelIndex   =   0
      Top             =   39
      VerticalCenter  =   0
      Visible         =   True
      Width           =   32
      ZIndex          =   1
      _DeclareLineRendered=   False
      _HorizontalPercent=   0.0
      _IsEmbedded     =   False
      _Locked         =   False
      _NeedsRendering =   True
      _OfficialControl=   False
      _OpenEventFired =   False
      _VerticalPercent=   0.0
   End
   Begin WebButton btnNext
      Caption         =   "Next>"
      Cursor          =   0
      Enabled         =   True
      Height          =   22
      HelpTag         =   ""
      HorizontalCenter=   0
      Index           =   -2147483648
      Left            =   726
      LockBottom      =   False
      LockedInPosition=   False
      LockHorizontal  =   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      LockVertical    =   False
      Scope           =   0
      Style           =   "0"
      TabOrder        =   10
      Top             =   473
      VerticalCenter  =   0
      Visible         =   True
      Width           =   100
      ZIndex          =   1
      _DeclareLineRendered=   False
      _HorizontalPercent=   0.0
      _IsEmbedded     =   False
      _Locked         =   False
      _NeedsRendering =   True
      _OfficialControl=   False
      _OpenEventFired =   False
      _VerticalPercent=   0.0
   End
   Begin conCreditCard CreditCard
      Cursor          =   0
      Enabled         =   True
      Height          =   420
      HelpTag         =   ""
      HorizontalCenter=   0
      Index           =   -2147483648
      Left            =   20
      LockBottom      =   False
      LockedInPosition=   False
      LockHorizontal  =   True
      LockLeft        =   False
      LockRight       =   False
      LockTop         =   True
      LockVertical    =   False
      Scope           =   0
      ScrollbarsVisible=   0
      Style           =   "997821280"
      TabOrder        =   4
      Top             =   48
      VerticalCenter  =   0
      Visible         =   False
      Width           =   910
      ZIndex          =   1
      _DeclareLineRendered=   False
      _HorizontalPercent=   0.0
      _IsEmbedded     =   False
      _Locked         =   False
      _NeedsRendering =   True
      _OfficialControl=   False
      _OpenEventFired =   False
      _ShownEventFired=   False
      _VerticalPercent=   0.0
   End
   Begin conConfirmation Confirmation
      Cursor          =   0
      Enabled         =   True
      Height          =   420
      HelpTag         =   ""
      HorizontalCenter=   0
      Index           =   -2147483648
      Left            =   20
      LockBottom      =   False
      LockedInPosition=   False
      LockHorizontal  =   True
      LockLeft        =   False
      LockRight       =   False
      LockTop         =   True
      LockVertical    =   False
      Scope           =   0
      ScrollbarsVisible=   0
      Style           =   "997821280"
      TabOrder        =   5
      Top             =   48
      VerticalCenter  =   0
      Visible         =   False
      Width           =   910
      ZIndex          =   1
      _DeclareLineRendered=   False
      _HorizontalPercent=   0.0
      _IsEmbedded     =   False
      _Locked         =   False
      _NeedsRendering =   True
      _OfficialControl=   False
      _OpenEventFired =   False
      _ShownEventFired=   False
      _VerticalPercent=   0.0
   End
   Begin conProcessing Processing
      Cursor          =   0
      Enabled         =   True
      Height          =   420
      HelpTag         =   ""
      HorizontalCenter=   0
      Index           =   -2147483648
      Left            =   20
      LockBottom      =   False
      LockedInPosition=   False
      LockHorizontal  =   True
      LockLeft        =   False
      LockRight       =   False
      LockTop         =   True
      LockVertical    =   False
      Scope           =   0
      ScrollbarsVisible=   0
      Style           =   "997821280"
      TabOrder        =   6
      Top             =   48
      VerticalCenter  =   0
      Visible         =   False
      Width           =   910
      ZIndex          =   1
      _DeclareLineRendered=   False
      _HorizontalPercent=   0.0
      _IsEmbedded     =   False
      _Locked         =   False
      _NeedsRendering =   True
      _OfficialControl=   False
      _OpenEventFired =   False
      _ShownEventFired=   False
      _VerticalPercent=   0.0
   End
   Begin conMemInfo MemInfo
      Cursor          =   0
      Enabled         =   True
      Height          =   420
      HelpTag         =   ""
      HorizontalCenter=   0
      Index           =   -2147483648
      Left            =   20
      lnErrorCount    =   -1
      LockBottom      =   False
      LockedInPosition=   False
      LockHorizontal  =   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      LockVertical    =   False
      mbIsCPD         =   False
      msEmail         =   ""
      msNameSuffix    =   ""
      msRegion        =   ""
      Scope           =   0
      ScrollbarsVisible=   0
      Style           =   "997821280"
      TabOrder        =   12
      Top             =   48
      VerticalCenter  =   0
      Visible         =   True
      Width           =   910
      ZIndex          =   1
      _DeclareLineRendered=   False
      _HorizontalPercent=   0.0
      _IsEmbedded     =   False
      _Locked         =   False
      _NeedsRendering =   True
      _OfficialControl=   False
      _OpenEventFired =   False
      _ShownEventFired=   False
      _VerticalPercent=   0.0
   End
   Begin WebLabel lblRecNo
      Cursor          =   1
      Enabled         =   True
      HasFocusRing    =   True
      Height          =   22
      HelpTag         =   ""
      HorizontalCenter=   0
      Index           =   -2147483648
      Left            =   20
      LockBottom      =   False
      LockedInPosition=   False
      LockHorizontal  =   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      LockVertical    =   False
      Multiline       =   False
      Scope           =   0
      Style           =   "-1"
      TabOrder        =   7
      Text            =   ""
      Top             =   471
      VerticalCenter  =   0
      Visible         =   True
      Width           =   100
      ZIndex          =   1
      _DeclareLineRendered=   False
      _HorizontalPercent=   0.0
      _IsEmbedded     =   False
      _Locked         =   False
      _NeedsRendering =   True
      _OfficialControl=   False
      _OpenEventFired =   False
      _VerticalPercent=   0.0
   End
   Begin WebLabel lblVersion
      Cursor          =   1
      Enabled         =   True
      HasFocusRing    =   True
      Height          =   22
      HelpTag         =   ""
      HorizontalCenter=   0
      Index           =   -2147483648
      Left            =   777
      LockBottom      =   False
      LockedInPosition=   False
      LockHorizontal  =   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      LockVertical    =   False
      Multiline       =   False
      Scope           =   0
      Style           =   "219244006"
      TabOrder        =   13
      Text            =   "Version:"
      Top             =   24
      VerticalCenter  =   0
      Visible         =   True
      Width           =   137
      ZIndex          =   1
      _DeclareLineRendered=   False
      _HorizontalPercent=   0.0
      _IsEmbedded     =   False
      _Locked         =   False
      _NeedsRendering =   True
      _OfficialControl=   False
      _OpenEventFired =   False
      _VerticalPercent=   0.0
   End
End
#tag EndWebPage

#tag WindowCode
	#tag Event
		Sub Open()
		  'if AppmbAffiliateGove then
		  'MemType.popType.Text = "Affiliate"
		  'end
		  
		End Sub
	#tag EndEvent

	#tag Event
		Sub Shown()
		  btnNext.Visible = True
		  lblVersion.Text = "Version: " + Str(app.MajorVersion) + "." + Str(app.MinorVersion) + "." + Str(app.BugVersion ) + " Build: "  + Str(app.NonReleaseVersion)
		  
		End Sub
	#tag EndEvent


	#tag Method, Flags = &h0
		Function CreateMsg() As String
		  
		  Dim rs, rsA as RecordSet
		  Dim lsMsg as String
		  
		  
		  rs = Session.sesWebDB.SQLSelect("Select * from memapplications where memappkwy = " + Str(Session.gnRecNo))
		  
		  if Session.sesWebDB.CheckDBError("Error sending Email Notification, contact support@aspe.org") then Return ""
		  
		  rsa = Session.sesAspeDB.SQLSelect("Select * from tblpeople where PersonID = " + Str(Session.gnPersonID))
		  
		  if Session.sesAspeDB.CheckDBError("Error sending Email Notification, contact support@aspe.org") then Return ""
		  
		  lsMsg = "<body><table border=""1"" cellspacing=""2"" cellpadding=""2"">"
		  
		  lsMsg = lsMsg +  "<tr><td width=""128"">Application ID: </td><td><strong>"
		  lsMsg = lsMsg +  rs.Field("memappkwy").StringValue
		  lsMsg = lsMsg +  "</strong></td></tr>\n"
		  
		  lsMsg = lsMsg +  "<tr><td width=""128"">Authorize.NET Transaction ID: </td><td><strong>"
		  lsMsg = lsMsg +  rs.Field("TransactionID").StringValue
		  lsMsg = lsMsg +  "</strong></td></tr>\n"
		  lsMsg = lsMsg +  "<tr><td width=""128"">AmountUSD: $</td><td><strong>"
		  lsMsg = lsMsg +  CreditCard.lblGrandTotal.Text
		  lsMsg = lsMsg +  "</strong></td></tr>\n"
		  lsMsg = lsMsg +  "<tr><td width=""128""> <hr /></td><td><strong> <hr />"
		  lsMsg = lsMsg +  "</strong></td></tr>\n"
		  
		  lsMsg = lsMsg +  "<tr><td width=""128"">Billing Address: </td><td><strong>"
		  lsMsg = lsMsg +  rs.Field("BillingAddress").StringValue
		  lsMsg = lsMsg +  "</strong></td></tr>\n"
		  
		  lsMsg = lsMsg +  "<tr><td width=""128"">City: </td><td><strong>"
		  lsMsg = lsMsg +  rs.Field("BillingCity").StringValue
		  lsMsg = lsMsg +  "</strong></td></tr>\n"
		  
		  lsMsg = lsMsg +  "<tr><td width=""128"">State or Province: </td><td><strong>"
		  lsMsg = lsMsg +  rs.Field("BillingStateProvince").StringValue
		  lsMsg = lsMsg +  "</strong></td></tr>\n"
		  
		  lsMsg = lsMsg +  "<tr><td width=""128"">Zip: </td><td><strong>"
		  lsMsg = lsMsg +  rs.Field("BillingZip").StringValue
		  lsMsg = lsMsg +  "</strong></td></tr>\n"
		  
		  lsMsg = lsMsg +  "<tr><td width=""128"">Country: </td><td><strong>"
		  lsMsg = lsMsg +  rs.Field("BillingCountry").StringValue
		  lsMsg = lsMsg +  "</strong></td></tr>\n"
		  
		  lsMsg = lsMsg +  "<tr><td width=""128"">Postal Code: </td><td><strong>"
		  lsMsg = lsMsg +  rs.Field("BillingPostalCode").StringValue
		  lsMsg = lsMsg +  "</strong></td></tr>\n"
		  
		  lsMsg = lsMsg +  "<tr><td width=""128"">PhoneDay: </td><td><strong>"
		  lsMsg = lsMsg +  rs.Field("BillingPhoneDay").StringValue
		  lsMsg = lsMsg +  "</strong></td></tr>\n"
		  
		  lsMsg = lsMsg +  "<tr><td width=""128"">Fax: </td><td><strong>"
		  lsMsg = lsMsg +  rs.Field("BillingFax").StringValue
		  lsMsg = lsMsg +  "</strong></td></tr>\n"
		  
		  lsMsg = lsMsg +  "<tr><td width=""128"">E-Mail: </td><td><strong>"
		  lsMsg = lsMsg +  rs.Field("BillingEmail").StringValue
		  lsMsg = lsMsg +  "</strong></td></tr>\n"
		  
		  lsMsg = lsMsg +  "<tr><td width=""128"">Card Holders First Name: </td><td><strong>"
		  lsMsg = lsMsg +  rs.Field("CardHolderFName").StringValue
		  lsMsg = lsMsg +  "</strong></td></tr>\n"
		  
		  lsMsg = lsMsg +  "<tr><td width=""128"">Card Holders Last Name: </td><td><strong>"
		  lsMsg = lsMsg +  rs.Field("CardHolderLName").StringValue
		  lsMsg = lsMsg +  "</strong></td></tr>\n"
		  
		  
		  lsMsg = lsMsg +  "<tr><td width=""128"">SPECIAL REINSTATEMENT:</td><td><strong>"
		  lsMsg = lsMsg +  "Membership Number: " + rs.Field("MembershipNum").StringValue
		  lsMsg = lsMsg +  "</strong></td></tr>\n"
		  
		  
		  lsMsg = lsMsg +  "<tr><td width=""128""><strong>Reinstatement Application"
		  lsMsg = lsMsg +  "</strong> </td></tr><td>All information has been updated.</td>"
		  
		  
		  
		  
		  Dim lsDBFormat as String
		  Select Case rs.Field("DataBookFormat").StringValue
		  Case "CD"
		    lsDBFormat = "CD-ROM"
		    
		  Case "SC"
		    lsDBFormat = "Soft Cover"
		    
		  Case "BO"
		    lsDBFormat = "Both Soft Cover and CD-ROM"
		    
		  end
		  
		  lsMsg = lsMsg +  "<tr><td width=""128"">Plumbing Engineering Design Hadbook Format: </td><td><strong>"
		  lsMsg = lsMsg +  lsDBFormat
		  lsMsg = lsMsg +  "</strong></td></tr>\n"
		  
		  
		  
		  lsMsg = lsMsg +  "<tr><td width=""128"">Name: </td><td><strong>"
		  lsMsg = lsMsg +  rsa.Field("NamePrefix").StringValue + " " + rsA.Field("FirstName").StringValue + " " + rsA.Field("Middle").StringValue  + " " + rsA.Field("LastName").StringValue + " " + rsA.Field("NameSuffix").StringValue
		  lsMsg = lsMsg +  "</strong></td></tr>\n"
		  
		  lsMsg = lsMsg +  "<tr><td width=""128"">Primary Email: </td><td><strong>"
		  lsMsg = lsMsg +  rsa.Field("Email").StringValue
		  lsMsg = lsMsg +  "</strong></td></tr>\n"
		  
		  
		  lsMsg = lsMsg +  "<hr>"
		  lsMsg = lsMsg +  "<br>"
		  
		  'lsMsg = lsMsg +  " <table width=""40%"" border=""1"" align=""left"" cellpadding=""1"" cellspacing=""1"">"
		  'lsMsg = lsMsg +  "<tr>"
		  'lsMsg = lsMsg +  "<td colspan=""3""><font size=""2"" face=""Arial,Helvetica,Geneva,Swiss,SunSans-Regular""><b>Registration (P.E.) </b></font></td>"
		  'lsMsg = lsMsg +  "</tr>\n"
		  'lsMsg = lsMsg +  "<tr>"
		  'lsMsg = lsMsg +  "<td><font size=""1"" face=""Arial,Helvetica,Geneva,Swiss,SunSans-Regular"">State</font></td>"
		  'lsMsg = lsMsg +  "<td><font size=""1"" face=""Arial,Helvetica,Geneva,Swiss,SunSans-Regular"">Certificate No. </font></td>"
		  'lsMsg = lsMsg +  "<td><font size=""1"" face=""Arial,Helvetica,Geneva,Swiss,SunSans-Regular"">Branch</font></td>"
		  'lsMsg = lsMsg +  "</tr>\n"
		  'lsMsg = lsMsg +  "<tr>"
		  'lsMsg = lsMsg +  "<td>$_POST[regState]</td>"
		  'lsMsg = lsMsg +  "<td>$_POST[regCertificate]</td>"
		  'lsMsg = lsMsg +  "<td>$_POST[regBranch]</td>"
		  'lsMsg = lsMsg +  "</tr>\n"
		  'lsMsg = lsMsg +  "</table>"
		  
		  lsMsg = lsMsg +  "<hr>"
		  'lsMsg = lsMsg +  "<br>"
		  
		  lsMsg = lsMsg +  "</BODY>"
		  
		  Return lsMsg
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub MoveToPage(Skip as Boolean = False)
		  If Skip then return
		  MemInfo.Visible = False
		  CreditCard.Visible = False
		  Confirmation.Visible = False
		  Processing.Visible = False
		  
		  btnNext.Visible = True
		  btnPrevious.Visible = True
		  
		  Select case msCurrentScreen
		    
		  Case "MemInfo"   '1
		    MemInfo.Visible = True
		    btnPrevious.Visible = False
		    Meminfo.popNamePrefix.SetFocus
		    
		    
		    
		  Case "CreditCard"
		    CreditCard.Visible = True
		    CreditCard.txtCCNumber.SetFocus
		    btnNext.Caption = "Next>"
		    
		    
		  Case "Confirmation"
		    Confirmation.txtCCCity.Text = CreditCard.txtCCCity.Text
		    Confirmation.txtCCCountry.Text = CreditCard.txtCCCountry.Text
		    Confirmation.txtCCEmail.Text = CreditCard.txtCCEmail.Text
		    Confirmation.txtCCFirst.Text = CreditCard.txtCCFirst.Text
		    Confirmation.txtCCLast.Text = CreditCard.txtCCLast.Text
		    Confirmation.txtCCNumber.Text = CreditCard.txtCCNumber.Text
		    Confirmation.txtCCPhoneHome.Text = CreditCard.txtCCPhoneHome.Text
		    Confirmation.txtCCState.Text = CreditCard.txtCCState.Text
		    Confirmation.txtCCStreetAddr.Text = CreditCard.txtCCStreetAddr.Text
		    Confirmation.txtCCZip.Text = CreditCard.txtCCZip.Text
		    Confirmation.txtCVV.Text = CreditCard.txtCVV.Text
		    Confirmation.txtExpMonth.Text = CreditCard.txtExpMonth.Text
		    Confirmation.txtExpYear.Text = CreditCard.txtExpYear.Text
		    
		    Confirmation.lblDBChoice.Text = CreditCard.lblDBChoice.Text
		    Confirmation.lblDatabook.Text = CreditCard.lblDBFormat.Text
		    Confirmation.lblMemshipCost.Text = CreditCard.lblMemshipCost.Text
		    Confirmation.lblGrandTotal.Text = CreditCard.lblGrandTotal.Text
		    Confirmation.Visible= True
		    btnNext.Caption = "Submit Order"
		    
		  Case "Processing"
		    Processing.Visible = True
		    gbCCDone = False
		    btnNext.Enabled = False
		    prgProgress.Visible = True
		    
		    System.DebugLog("N4 Timer")
		    Timer1.Mode = Timer.ModeSingle
		    System.DebugLog("After timer")
		    
		    ''btnNext.Enabled = False
		    'btnNext.Caption = "Next"
		    
		  end
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub QuitTimer()
		  'TimerQuit.Mode = Timer.ModeSingle
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SendApplication()
		  Dim rs as RecordSet
		  Dim Msg as New EmailMessage
		  
		  System.DebugLog("In Send Application")
		  
		  '
		  rs = Session.sesWebDB.SQLSelect("Select * from memapplications where memappkwy = " + Str(Session.gnRecNo))
		  
		  'AddHandler SMTPServer.
		  
		  
		  SMTPServer = New SMTPServer1
		  SMTPServer.Address = "localhost" 'csBulkMailSMTPServer
		  SMTPServer.Port = 25 'cnBulkEmailPort
		  SMTPServer.Username = "Hawkcode@gmail.com " '""   'csBulkMailSMTPUserID
		  SMTPServer.Password = ""   'csBulkEmailSMTPPassword
		  
		  ' Not needed when deployed
		  'SMTPServer.ConnectionType = SMTPSecureSocket.SSLv2
		  'SMTPServer.Secure = True
		  'SMTPServer.Connect
		  
		  System.DebugLog("Last Error: " + Str(SMTPServer.LastErrorCode) )
		  
		  Msg.AddRecipient csEmailAddress
		  Msg.AddCCRecipient msUserEmail
		  
		  System.DebugLog("Sending email to: " + csEmailAddress)
		  Msg.subject = "Reinstatement Form - " + rs.Field("lastName").StringValue + ", " + rs.Field("firstName").StringValue + " " + rs.Field("middleName").StringValue
		  Msg.BodyHTML = CreateMsg()
		  If Msg.BodyHTML = "" then Return
		  
		  SMTPServer.Messages.Append( Msg)
		  SMTPServer.SendMail
		  
		  System.DebugLog("Last Error: " + Str(SMTPServer.LastErrorCode) )
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SendAuthorizeNet()
		  
		  dicAuth = New Dictionary
		  
		  
		  's.SetFormData(d)
		  '
		  'Dim lsStr as String
		  '
		  'lsStr = s.Post("http://aspe.org/xo/test/test.php", 10)
		  '
		  'lsStr = DefineEncoding(lsStr, Encodings.UTF8)
		  
		  Dim lsDesc as String
		  lsDesc = "Membership Reinstatement"
		  
		  
		  dicAuth.Value("x_email") = CreditCard.txtCCEmail.Text
		  dicAuth.Value("x_card_num") = CreditCard.txtCCNumber.Text
		  dicAuth.Value("x_exp_date") = CreditCard.txtExpMonth.Text + CreditCard.txtExpYear.Text
		  dicAuth.Value("x_card_code") = CreditCard.txtCVV.Text
		  dicAuth.Value("x_description") = lsDesc
		  dicAuth.Value("x_amount") = CreditCard.lblGrandTotal.Text
		  dicAuth.Value("x_first_name") = CreditCard.txtCCFirst.Text
		  dicAuth.Value("x_last_name") = CreditCard.txtCCLast.Text
		  dicAuth.Value("x_address") = CreditCard.txtCCStreetAddr.Text
		  dicAuth.Value("x_city") = CreditCard.txtCCCity.Text
		  dicAuth.Value("x_state") = CreditCard.txtCCState.Text
		  dicAuth.Value("x_zip") = CreditCard.txtCCZip.Text
		  dicAuth.Value("x_phone") = CreditCard.txtCCPhoneHome.Text
		  dicAuth.Value("x_country") = CreditCard.txtCCCountry.Text
		  
		  
		  Dim dicResultCode as New Dictionary
		  '
		  if CreditCard.txtCCNumber.Text = "11111111111111110" and CreditCard.txtCVV.Text = "9999" then
		    dicResultCode.Value("ResponseCode") = "Approved"
		    dicResultCode.Value("TransActionID") = "1234567890"
		    dicResultCode.Value("AVSResponse") = "All Match"
		  else
		    dicResultCode = ProcessCC(dicAuth, False )
		  end
		  
		  'Moved to CCDone
		  '
		  if dicResultCode.Value("ResponseCode")  = "Approved" then
		    Processing.txtResult.Text = "Result: Transaction Approved"+ EndOfLine
		    Processing.txtResult.Text =  Processing.txtResult.Text + "TransActionID: " + dicResultCode.Value("TransActionID")+ EndOfLine
		    Processing.txtResult.Text =  Processing.txtResult.Text + dicResultCode.Value("AVSResponse")+ EndOfLine
		    btnPrevious.Enabled = False
		    btnNext.Visible = False
		    if not UpdateTransaction( dicResultCode.Value("TransActionID"), "Approved") then
		      MsgBox("Error Unable to Send Application, Your transaction did go through though.")
		      Return
		    end
		    SendApplication
		  else
		    Processing.txtResult.Text = "Result: "+ dicResultCode.Value("ResponseCode") + EndOfLine
		    Processing.txtResult.Text =  Processing.txtResult.Text +  "--- " + dicResultCode.Value("ResponseReasonCode") + EndOfLine + EndOfLine
		    Processing.txtResult.Text =  Processing.txtResult.Text +  dicResultCode.Value("ResponseReasonText") + EndOfLine + EndOfLine
		    if dicResultCode.HasKey("AVSResponse") then
		      Processing.txtResult.Text =  Processing.txtResult.Text +  dicResultCode.Value("AVSResponse") + EndOfLine
		    end
		    btnPrevious.Enabled = True
		  end
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetNextPage(bNext as Boolean)
		  'prgProgressVisible = True
		  'if bNext then btnNext.enabled = False
		  
		  
		  Select case msCurrentScreen
		    
		  Case "MemInfo"   '1
		    if not MemInfo.ValidateAll then return
		    if not Meminfo.SaveMemInfo then return
		    msCurrentScreen = "CreditCard"
		    
		    
		    
		    
		  Case "CreditCard"
		    if bNext then
		      msCurrentScreen = "Confirmation"
		      if not CreditCard.ValidateAll then return
		      If Not CreditCard.SaveCC then Return
		    else
		      msCurrentScreen = "MemInfo"
		    end
		    
		    
		  Case "Confirmation"
		    if bNext then
		      msCurrentScreen = "Processing"
		      
		    else
		      msCurrentScreen = "CreditCard"
		    end
		    
		  Case "Processing"
		    if bNext then
		      
		    else
		      msCurrentScreen = "CreditCard"
		    end
		  end
		  
		  MoveToPage
		  if bNext and msCurrentScreen <>   "Processing" then btnNext.enabled = True
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function UpdateTransaction(lsTransID as String, lsResultCode as String) As Boolean
		  
		  Dim lsStr as String
		  Dim lnLen as Integer
		  Dim oSQL as new cSmartSQL
		  Dim rs as RecordSet
		  
		  
		  'Me.MouseCursor = System.Cursors.Wait
		  
		  oSQL.StatementType = eStatementType.Type_Update
		  
		  oSql.ClearFields
		  oSQL.ClearValues
		  
		  oSQL.AddTable "memapplications"
		  
		  Dim lsExpDate as String
		  
		  
		  oSQL.AddFields "TransactionID",     "ResultCode"
		  oSQL.AddValues lsTransID,          lsResultCode
		  
		  oSQL.AddSimpleWhereClause "memappkwy", Session.gnRecNo
		  
		  Session.sesWebDB.SQLExecute(oSQL.SQL)
		  
		  
		  
		  if Session.sesWebDB.CheckDBError then
		    MsgBox(Session.sesWebDB.ErrorMessage)
		    Return False
		  end
		  
		  Return True
		  
		  
		  
		End Function
	#tag EndMethod


	#tag Note, Name = Untitled
		
		
		SMTPServer As SMTPSecureSocket
		
		In my Project.App.Open I initialise this object
		
		SMTPServer = New SMTPSecureSocket
		t
		heMailServer.Address = "smtp.gmail.com
		
		SMTPServer.Port = 465
		
		SMTPServer.Username = <your gmail username>
		
		SMTPServer.Password = "<your gmail password>"
		
		SMTPServer.ConnectionType = SMTPSecureSocket.SSLv23
		
		SMTPServer.Secure = True
		
		It’s important that this is instantiated in the App object as we’ll be using asynchronous communication.
		
		In the Action event of the submit button I will send the message.
		
		Dim msg As New EmailMessage
		
		msg.AddRecipient “<my address>”
		
		msg.subject = Subject.Text
		
		msg.BodyPlainText = Message.Text
		
		App.SMTPServer.AppendMessage msg
		
		App.SMTPServer.SendMessage
		
		
	#tag EndNote


	#tag Property, Flags = &h0
		dicAuth As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h0
		gbCCDone As Boolean = True
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			'MemInfo
		#tag EndNote
		msCurrentScreen As String = "MemInfo"
	#tag EndProperty

	#tag Property, Flags = &h0
		msUserEmail As String
	#tag EndProperty

	#tag Property, Flags = &h0
		SMTPServer As SMTPSecureSocket
	#tag EndProperty


	#tag Constant, Name = cnBulkEmailPort, Type = Double, Dynamic = False, Default = \"465", Scope = Public
	#tag EndConstant

	#tag Constant, Name = csBulkEmailSMTPPassword, Type = String, Dynamic = False, Default = \"aspe8614", Scope = Public
	#tag EndConstant

	#tag Constant, Name = csBulkMailSMTPServer, Type = String, Dynamic = False, Default = \"smtp.gmail.com", Scope = Public
	#tag EndConstant

	#tag Constant, Name = csBulkMailSMTPUserID, Type = String, Dynamic = False, Default = \"aspenationalrich@gmail.com", Scope = Public
	#tag EndConstant

	#tag Constant, Name = csEmailAddress, Type = String, Dynamic = False, Default = \"Admin@aspe.org", Scope = Public
	#tag EndConstant


#tag EndWindowCode

#tag Events btnPrevious
	#tag Event
		Sub Action()
		  SetNextPage(False)
		  
		  'HTMLViewer1.LoadPage(lsStr)
		  
		  'me.ExecuteJavaScript("window.open('../test/test.php?Name=Tom&Title=Dick','_self');")
		  'me.ExecuteJavaScript("alert('Hello!');")
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events Timer1
	#tag Event
		Sub Action()
		  
		  btnNext.Enabled = False
		  System.DebugLog("Timer Start")
		  SendAuthorizeNet
		  prgProgress.Visible = False
		  System.DebugLog("Timer end")
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events SMTPServer1
	#tag Event
		Sub ServerError(ErrorID as integer, ErrorMessage as string, Email as EmailMessage)
		  System.DebugLog("Server Error: " + ErrorMessage)
		End Sub
	#tag EndEvent
	#tag Event
		Sub MessageSent(Email as EmailMessage)
		  System.DebugLog("Mail sent")
		End Sub
	#tag EndEvent
	#tag Event
		Sub ConnectionEstablished(greeting as string)
		  System.DebugLog("Connection Made")
		End Sub
	#tag EndEvent
	#tag Event
		Sub Error()
		  System.DebugLog("Mail Error")
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events dlgCorrectPersonYN
	#tag Event
		Sub Open()
		  
		End Sub
	#tag EndEvent
	#tag Event
		Sub Dismissed()
		  if Me.SelectedButton.Caption = "No" then
		    frmAppllcation.MemInfo.txtMembershipNumber.Text = ""
		  end
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events TimerQuit
	#tag Event
		Sub Action()
		  App.Quit
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnNext
	#tag Event
		Sub Action()
		  'me.Enabled = False
		  SetNextPage(True)
		  
		  'HTMLViewer1.LoadPage(lsStr)
		  
		  'me.ExecuteJavaScript("window.open('../test/test.php?Name=Tom&Title=Dick','_self');")
		  'me.ExecuteJavaScript("alert('Hello!');")
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag ViewBehavior
	#tag ViewProperty
		Name="Cursor"
		Visible=true
		Group="Behavior"
		InitialValue="0"
		Type="Integer"
		EditorType="Enum"
		#tag EnumValues
			"0 - Auto"
			"1 - Standard Pointer"
			"2 - Finger Pointer"
			"3 - IBeam"
			"4 - Wait"
			"5 - Help"
			"6 - Arrow All Directions"
			"7 - Arrow North"
			"8 - Arrow South"
			"9 - Arrow East"
			"10 - Arrow West"
			"11 - Arrow North East"
			"12 - Arrow North West"
			"13 - Arrow South East"
			"14 - Arrow South West"
			"15 - Splitter East West"
			"16 - Splitter North South"
			"17 - Progress"
			"18 - No Drop"
			"19 - Not Allowed"
			"20 - Vertical IBeam"
			"21 - Crosshair"
		#tag EndEnumValues
	#tag EndViewProperty
	#tag ViewProperty
		Name="Enabled"
		Group="Behavior"
		InitialValue="True"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Height"
		Visible=true
		Group="Behavior"
		InitialValue="400"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="HelpTag"
		Visible=true
		Group="Behavior"
		Type="String"
		EditorType="MultiLineEditor"
	#tag EndViewProperty
	#tag ViewProperty
		Name="HorizontalCenter"
		Group="Behavior"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Index"
		Group="ID"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="IsImplicitInstance"
		Group="Behavior"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Left"
		Group="Position"
		InitialValue="0"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="LockBottom"
		Group="Behavior"
		InitialValue="False"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="LockHorizontal"
		Group="Behavior"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="LockLeft"
		Group="Behavior"
		InitialValue="False"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="LockRight"
		Group="Behavior"
		InitialValue="False"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="LockTop"
		Group="Behavior"
		InitialValue="False"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="LockVertical"
		Group="Behavior"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinHeight"
		Visible=true
		Group="Behavior"
		InitialValue="400"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinWidth"
		Visible=true
		Group="Behavior"
		InitialValue="600"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Name"
		Visible=true
		Group="ID"
		Type="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Super"
		Visible=true
		Group="ID"
		Type="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="TabOrder"
		Group="Behavior"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Title"
		Visible=true
		Group="Behavior"
		InitialValue="Untitled"
		Type="String"
		EditorType="MultiLineEditor"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Top"
		Group="Position"
		InitialValue="0"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="VerticalCenter"
		Group="Behavior"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Visible"
		Group="Behavior"
		InitialValue="True"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Width"
		Visible=true
		Group="Behavior"
		InitialValue="600"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="ZIndex"
		Group="Behavior"
		InitialValue="1"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="_HorizontalPercent"
		Group="Behavior"
		Type="Double"
	#tag EndViewProperty
	#tag ViewProperty
		Name="_ImplicitInstance"
		Group="Behavior"
		InitialValue="False"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="_IsEmbedded"
		Group="Behavior"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="_Locked"
		Group="Behavior"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="_NeedsRendering"
		Group="Behavior"
		InitialValue="True"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="_OfficialControl"
		Group="Behavior"
		InitialValue="False"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="_ShownEventFired"
		Group="Behavior"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="_VerticalPercent"
		Group="Behavior"
		Type="Double"
	#tag EndViewProperty
#tag EndViewBehavior
