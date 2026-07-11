VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CParamPenteDZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CParamPenteDZ - Couple pente (%) + decalage fixe DZ, chacun activable
'
' Utilise trois fois par CMstSettings :
'   oDecalage : pente transversale + DZ fixe (interpolation)
'   oChemin   : chemin principal (ponctuelle)
'   oRayon    : rayonnement (ponctuelle)
'
' Les membres valent 0 / False par defaut (initialisation VBA), ce qui
' correspond aux valeurs de depart voulues : rien n'est applique.
'==============================================================================
Option Explicit

Public Pente        As Double    ' pente en %
Public PenteActive  As Boolean   ' True = appliquer la pente
Public DZ           As Double    ' decalage fixe en Z (u.m.)
Public DZActive     As Boolean   ' True = appliquer le decalage DZ
