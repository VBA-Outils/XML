VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IndenterXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Indenter un flux XML - VBA
'
' @Class IndenterXML
' @author vincent.rosset@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'
' Copyright (c) 2023, Vincent ROSSET
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'     * Redistributions of source code must retain the above copyright
'       notice, this list of conditions and the following disclaimer.
'     * Redistributions in binary form must reproduce the above copyright
'       notice, this list of conditions and the following disclaimer in the
'       documentation and/or other materials provided with the distribution.
'     * Neither the name of the <organization> nor the
'       names of its contributors may be used to endorse or promote products
'       derived from this software without specific prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
' ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
' DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
' (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
' ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

' Nécessite d'activer la référence "Microsoft Scripting RunTime"
'
' Dans l'éditeur de macros (Alt+F11): Menu Outils \ Références
' Cochez les lignes :
' -- "Microsoft XML, v3"
' Cliquez sur le bouton OK pour valider.

Option Explicit

'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'
' Propriétés de la classe :
' -------------------------
'   DocumentXML                        : Document XML
'   NomFichier                         : nom du fichier (y compris le chemin d'accès)
'
' Méthodes de la classe :
' -----------------------
'   EnregistrerSous                    : Enregistrer sous le nom de fichier le contenu de l'objet ADODB.Stream
'   IndenterDocumentXML                : Indenter un document XML et enregistrer l'indentation dans un fichier
'
'-----------------------------------------------------------------------------------------------------------------------------------------------------------

Private poADODB As ADODB
' Objet ADODB.stream
Private poFichier As Object
' Nom du fichier sous lequel est enregistré le flux XML
Private psNomFichier As String

'---------------------------------------------------------------------------------------------------
' Document XML
'---------------------------------------------------------------------------------------------------
Private poDocumentXML As MSXML2.DOMDocument

' *---------------------------------------------------------------------------------------------------*
' * Propriété : Document XML                                                                          *
' *---------------------------------------------------------------------------------------------------*
Property Set DocumentXML(ByRef oDocumentXML As MSXML2.DOMDocument)
    Set poDocumentXML = oDocumentXML
End Property

' *---------------------------------------------------------------------------------------------------*
' * Propriété : nom du fichier                                                                        *
' *---------------------------------------------------------------------------------------------------*
Property Let NomFichier(sNomFichier As String)

    If Trim(sNomFichier) = "" Then
        Err.Raise vbObjectError + 9, , "Le nom du fichier doit être renseigné."
    Else
        psNomFichier = Trim(sNomFichier)
    End If

End Property

Property Get NomFichier() As String
    NomFichier = psNomFichier
End Property

' *---------------------------------------------------------------------------------------------------*
' * Indenter le flux XML pour la lisibilité avec XMLReader et XMLWriter                               *
' *---------------------------------------------------------------------------------------------------*
' * Propriété utilisée :                                                                              *
' *   -- DocumentXML         (Obligatoire)                                                            *
' *---------------------------------------------------------------------------------------------------*
Public Sub IndenterDocumentXML(sEncodage As String)

    ' Parser XML
    Dim oXMLReader As Object, oXmlWriter As Object
    
    ' Vérifie si le document XML a été défini
    If poDocumentXML Is Nothing Then
        Err.Raise vbObjectError + 61, , "Le document XML n'existe pas."
    End If
    
    Set oXMLReader = CreateObject("MSXML2.SAXXMLReader")
    Set oXmlWriter = CreateObject("MSXML2.MXXMLWriter")
    
    ' Fichier de type Texte
    Set poADODB.Fichier = poFichier
    poADODB.TypeFichier = FICHIER_TEXTE
    poADODB.EncodageTxt = sEncodage
    poADODB.TypeAcces = ACCES_ECRITURE
    ' Ouverture du fichier
    poADODB.Ouvrir

    ' Forcer l'indentation
    oXmlWriter.indent = True
    oXmlWriter.Encoding = sEncodage
    ' Ecriture dans le flux ADODB
    oXmlWriter.output = poFichier
    Set oXMLReader.contentHandler = oXmlWriter
    Set oXMLReader.errorHandler = oXmlWriter
    ' Parse de l'objet Document
    oXMLReader.Parse poDocumentXML
    oXmlWriter.flush
    
    Set oXMLReader = Nothing
    Set oXmlWriter = Nothing

End Sub

' *---------------------------------------------------------------------------------------------------*
' * Définir le nom du fichier XML                                                                     *
' *---------------------------------------------------------------------------------------------------*
Public Sub SélectionnerNomFichier()
    
    With poADODB
        .Filtre = INDEX_FILTRE_XML
        .SelectionnerFichierEnregistrerSous
        ' Enregistrer le résultat dans le fichier "NomFichier"
        psNomFichier = .NomFichier
    End With
    
End Sub

' *---------------------------------------------------------------------------------------------------*
' * Enregistrer le flux XML dans un fichier                                                           *
' *---------------------------------------------------------------------------------------------------*
' * Propriétés utilisées :                                                                            *
' *   -- nomFichier          (Obligatoire)                                                            *
' *   -- DocumentXML         (Obligatoire)                                                            *
' *---------------------------------------------------------------------------------------------------*
Public Sub EnregistrerSous()
    
    If Trim(psNomFichier) = "" Then
        Err.Raise vbObjectError + 9, , "Le nom du fichier doit être renseigné."
    Else
        poADODB.EnregistrerSous
    End If
    
End Sub

Private Sub Class_Initialize()
    
    ' Création d'un objet ADODB
    Set poADODB = New ADODB
    Set poFichier = CreateObject("ADODB.STREAM")
       
End Sub

Private Sub Class_Terminate()
    Set poADODB = Nothing
    Set poFichier = Nothing
End Sub
