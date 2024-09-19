Option Explicit

Public WithEvents lblGrupo As MSForms.Label

Private Sub lblGrupo_Click()
    lblGrupo.parent.Tag = lblGrupo.Tag
    lblGrupo.parent.Hide
End Sub

