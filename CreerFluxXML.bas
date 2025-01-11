Attribute VB_Name = "CreerFluxXML"
'
' Parser un flux XML
' https://github.com/VBA-Outils/XML
'
' @Module ParserFlux
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

' Nécessite d'activer les références "Microsoft Scripting RunTime" et "Microsoft XML. V3.0"
'
' Dans l'éditeur de macros (Alt+F11): Menu Outils \ Références
' Cochez les lignes "Microsoft Scripting RunTime" et "Microsoft XML. V3.0".
' Cliquez sur le bouton OK pour valider.

' Le module "Bibliotheque" et la classe "ADODB" doivent être importés.

Option Explicit
Option Compare Text

' *---------------------------------------------------------------------------------------------------*
' * Convertir un flux XML bouteille                                                                   *
' *---------------------------------------------------------------------------------------------------*
Public Sub ConvertirFluxXML()

    Dim sNomFichierEntree As String, oFichier As New ADODB
    
    ' Sélectionner le fichier Eau.xml
    With oFichier
        .Repertoire = Environ("OneDrive") & "\Documents\"
        If .RepertoireExiste = True Then .NomInitialFichier = .Repertoire
        .TitreBoiteDeDialogue = "Sélectionner le flux XML"
        .LibelleFiltre = "Fichier XML"
        .ExtensionFiltre = "*.xml"
        .SelectionnerFichier
        ' Si un nom de fichier a été sélectionné
        If .NomFichier <> "" Then
            ' Traiter le flux MXL
            Call TraiterFluxXML(.NomFichier)
        End If
    End With
    Set oFichier = Nothing
    
End Sub

' *---------------------------------------------------------------------------------------------------*
' * Traiter un flux XML                                                                               *
' *   -- Ajouter un noeud SOAP                                                                        *
' *   -- Ajouter un noeud Entête                                                                      *
' *   -- Ajouter un noeud Body                                                                        *
' *      -- Insérer dans le noeud Body le flux Eau.xml                                                *
' *   -- Indenter le flux XML généré                                                                  *
' *---------------------------------------------------------------------------------------------------*
Private Sub TraiterFluxXML(sNomFichierEntree As String)

    Dim oDocEntree As MSXML2.DOMDocument
    Dim oNoeudDoc As MSXML2.IXMLDOMNode, oNoeudEauE As MSXML2.IXMLDOMNode

    Dim oDocSortie As MSXML2.DOMDocument, oNoeudFrameWork As MSXML2.IXMLDOMNode, oNoeudHeaderS As MSXML2.IXMLDOMNode, oNoeudSoapS As MSXML2.IXMLDOMNode, oNoeudBodyS As MSXML2.IXMLDOMNode, oNoeudEauS As MSXML2.IXMLDOMNode
    Dim sNomfichierSortie As String
    
    Dim oFichierXML As New IndenterXML
    
    Call InitialiserTraitement

    ' Déclaration du DOM
    Set oDocEntree = New MSXML2.DOMDocument
    oDocEntree.async = False
    ' Chargement du flux XML
    oDocEntree.Load sNomFichierEntree
    
    If oDocEntree.parseError.ErrorCode <> 0 Then
        Call TerminerTraitement
        Set oDocEntree = Nothing
        Set oNoeudDoc = Nothing
        Set oDocSortie = Nothing
        Set oNoeudFrameWork = Nothing
        Set oNoeudHeaderS = Nothing
        Set oNoeudSoapS = Nothing
        Set oNoeudBodyS = Nothing
        Set oNoeudEauS = Nothing
        Set oFichierXML = Nothing
        Exit Sub
    End If
    
    ' Rechercher le noeud principal Eau
    For Each oNoeudDoc In oDocEntree.ChildNodes
        If oNoeudDoc.BaseName = "Eau" Then
            Set oNoeudEauE = oNoeudDoc
        End If
    Next oNoeudDoc
    
    ' Créer le DOM en sortie
    Set oDocSortie = New MSXML2.DOMDocument
    
    ' Ajouter le noeud Envelope
    Set oNoeudSoapS = oDocSortie.createElement("soap:Envelope")
    oDocSortie.appendChild oNoeudSoapS
    Call AjouterAttribut(oDocSortie, oNoeudSoapS, "xmlns:soap", "http://schemas.xmlsoap.org/soap/envelope/")
    
    ' Ajouter le noeud Header
    Call AjouterNoeud(oDocSortie, oNoeudSoapS, oNoeudHeaderS, "soap:Header")
    
    ' Ajouter le noeud FrameWork
    Call AjouterNoeud(oDocSortie, oNoeudHeaderS, oNoeudFrameWork, "FrameWork")
    Call AjouterAttribut(oDocSortie, oNoeudFrameWork, "xmlns", "http://frameWork.fr/")
    Call AjouterAttribut(oDocSortie, oNoeudFrameWork, "Attribut1", "ATT")
    
    ' Ajouter le noeud Body
    Call AjouterNoeud(oDocSortie, oNoeudSoapS, oNoeudBodyS, "soap:Body")
    
    ' Dans Body, ajouter le flux xml lu précédemment : Eau
    oNoeudBodyS.appendChild oNoeudEauE
    
    ' Appel de la classe d'indentation d'un flux XML avec écriture du résultat dans un fichier
    With oFichierXML
        Set .DocumentXML = oDocSortie
        .IndenterDocumentXML ("utf-8")
        .SélectionnerNomFichier
        .EnregistrerSous
        sNomfichierSortie = .NomFichier
    End With
    
    Set oFichierXML = Nothing
    Set oDocEntree = Nothing
    Set oNoeudDoc = Nothing
    Set oDocSortie = Nothing
    Set oNoeudFrameWork = Nothing
    Set oNoeudHeaderS = Nothing
    Set oNoeudSoapS = Nothing
    Set oNoeudBodyS = Nothing
    Set oNoeudEauS = Nothing
    
    Call TerminerTraitement
    
End Sub

' *---------------------------------------------------------------------------------------------------*
' * Ajout d'un noeud (enfant) à un noeud existant (parent)                                            *
' *---------------------------------------------------------------------------------------------------*
Private Sub AjouterNoeud(oDoc As MSXML2.DOMDocument, oNoeudParent As MSXML2.IXMLDOMNode, oNoeudEnfant As MSXML2.IXMLDOMNode, sNomNoeud As String, Optional sTexteNoeud As String)
    
    Set oNoeudEnfant = oDoc.createElement(sNomNoeud)
    If sTexteNoeud <> "" Then oNoeudEnfant.Text = sTexteNoeud
    oNoeudParent.appendChild oNoeudEnfant
    
End Sub

' *---------------------------------------------------------------------------------------------------*
' * Ajout d'un attribut à un noeud                                                                    *
' *---------------------------------------------------------------------------------------------------*
Private Sub AjouterAttribut(oDoc As MSXML2.DOMDocument, oNoeud As MSXML2.IXMLDOMNode, sNomAttribut As String, sTexteAttribut As String)

    Dim oAttribut As MSXML2.IXMLDOMNode
    
    Set oAttribut = oDoc.createAttribute(sNomAttribut)
    oAttribut.Text = sTexteAttribut
    oNoeud.Attributes.setNamedItem oAttribut
    Set oAttribut = Nothing

End Sub
