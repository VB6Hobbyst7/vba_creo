VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCreo_tbm 
   Caption         =   "Title Block Manager for Creo Parametric"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8040
   OleObjectBlob   =   "frmCreo_tbm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCreo_tbm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private crDrw As pfcls.IpfcDrawing

Private Sub CommandButton1_Click()
Unload Me
End Sub

Private Sub CommandButton2_Click()
Call writeToActiveDocument

Dim m2 As pfcls.IpfcModel2D
Set m2 = crDrw
m2.Regenerate

Dim ana As String
ana = Split(creoSsn.CurrentModel.filename, ".")(0)
AppActivate ana

Unload Me
End Sub

Private Sub CommandButton3_Click()
SaveSetting "Domisoft", "TBM_SE", "Default_Designer", designer.Text
End Sub

Private Sub CommandButton4_Click()
SaveSetting "Domisoft", "TBM_SE", "Default_Reviewer", reviewer.Text
End Sub

Private Sub CommandButton5_Click()
SaveSetting "Domisoft", "TBM_SE", "Default_Approver", approver.Text
End Sub
Private Sub designer_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
designer.Text = GetSetting("Domisoft", "TBM_SE", "Default_Designer", "")
End Sub
Private Sub reviewer_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
reviewer.Text = GetSetting("Domisoft", "TBM_SE", "Default_Reviewer", "")
End Sub
Private Sub approver_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
approver.Text = GetSetting("Domisoft", "TBM_SE", "Default_Approver", "")
End Sub

Private Sub UserForm_Initialize()

If creoSsn Is Nothing Then Call Conn2creo

Set crDrw = creoSsn.CurrentModel

Call getFromActive
Call loadDefault

stage.AddItem "E", 0
stage.AddItem "M", 1
stage.AddItem "P", 2

End Sub
Private Sub loadDefault()
If designer.Text = "" Then designer.Text = GetSetting("Domisoft", "TBM_SE", "Default_Designer", "")
If reviewer.Text = "" Then reviewer.Text = GetSetting("Domisoft", "TBM_SE", "Default_Reviewer", "")
If approver.Text = "" Then approver.Text = GetSetting("Domisoft", "TBM_SE", "Default_Approver", "")
If design_date.Text = "" Then design_date.Text = qDate(VBA.Date)
If review_date.Text = "" Then review_date.Text = qDate(VBA.Date)
If approve_date.Text = "" Then approve_date.Text = qDate(VBA.Date)
End Sub

Private Sub getFromActive()
name_cn.Text = getParam(crDrw, "PART_NAME_LANG")
model_no.Text = getParam(crDrw, "product_code")
material.Text = getParam(crDrw, "material_cn")
weight.Text = "&PRO_MP_MASS"
designer.Text = getParam(crDrw, "drafter")
design_date.Text = getParam(crDrw, "draft_date")
reviewer.Text = getParam(crDrw, "checker")
review_date.Text = getParam(crDrw, "check_date")
approver.Text = getParam(crDrw, "engr")
approve_date.Text = getParam(crDrw, "engr_date")
version.Text = getParam(crDrw, "qhc_version")
stage.Text = getParam(crDrw, "qhc_stage")
End Sub
Private Sub writeToActiveDocument()
Call editParam(crDrw, "PART_NAME_LANG", Trim(name_cn.Text))
Call editParam(crDrw, "product_code", Trim(model_no.Text))
Call editParam(crDrw, "material_cn", Trim(material.Text))
'Call editParam(crDrw, "material_cn", Trim(weight.Text))
Call editParam(crDrw, "drafter", Trim(designer.Text))
Call editParam(crDrw, "draft_date", Trim(design_date.Text))
Call editParam(crDrw, "checker", Trim(reviewer.Text))
Call editParam(crDrw, "check_date", Trim(review_date.Text))
Call editParam(crDrw, "engr", Trim(approver.Text))
Call editParam(crDrw, "engr_date", Trim(approve_date.Text))
Call editParam(crDrw, "qhc_version", Trim(version.Text))
Call editParam(crDrw, "qhc_stage", Trim(stage.Text))
End Sub




