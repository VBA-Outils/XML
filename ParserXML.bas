Attribute VB_Name = "ParserXML"
'
' Parser un flux XML
' https://github.com/VBA-Outils/XML
'
' @Module ParserXML
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

' Nécessite d'activer les références "Microsoft XML. V3.0"
'
' Dans l'éditeur de macros (Alt+F11): Menu Outils \ Références
' Cochez les lignes "Microsoft XML. V3.0".
' Cliquez sur le bouton OK pour valider.

Option Explicit
Option Compare Text

' Structure du flux XML utilisé :
' Eau
'  -> Bouteille
'      -> marque
'      -> Composition
'          -> ion
'          -> autre
'      -> source
'      -> code barre
'      -> contenance
'      -> pH

Type composition
    typeIon As String
    nomIon As String
End Type

Type listeComposition
    nbIons As Integer
    composition() As composition
End Type

Type source
    ville As String
    departement As String
End Type

Type Bouteille
    marque As String
    listeComposition As listeComposition
    source As source
    codeBarre As String
    contenance As Integer
    ph As Double
End Type

Type listeBouteille
    NbBouteilles As Integer
    Bouteille() As Bouteille
End Type

' *---------------------------------------------------------------------------------------------------*
' * Parser le flux                                                                                    *
' *---------------------------------------------------------------------------------------------------*
Public Function ParserFluxXML(Optional NomFichier As String, Optional FluxMXL As String) As listeBouteille

    Dim oDoc As MSXML2.DOMDocument, oNoeudDoc As MSXML2.IXMLDOMNode
    
    ' Si les 2 paramètres ne sont pas renseignés
    If NomFichier = "" And FluxMXL = "" Then Exit Function
    
    ' Déclaration du DOM
    Set oDoc = New MSXML2.DOMDocument
    ' Utiliser le mode synchrone
    oDoc.async = False
    ' Chargement du fichier ou flux XML en fonction du paramètre renseigné
    If NomFichier = "" Then
        oDoc.LoadXML FluxMXL
    Else
        oDoc.Load NomFichier
    End If
    
    ' Si aucune erreur lors du parse
    If oDoc.parseError.ErrorCode = 0 Then
        ' Parcourir tous les noeuds de la racine
        For Each oNoeudDoc In oDoc.ChildNodes
            ' 1er noeud : Eau
            If oNoeudDoc.BaseName = "Eau" Then
                ParserFluxXML = ParserEau(oNoeudDoc)
            End If
        Next oNoeudDoc
    End If
    
    Set oDoc = Nothing

End Function

' *---------------------------------------------------------------------------------------------------*
' * Parcourir tous les élements du noeud "Eau"                                                        *
' *---------------------------------------------------------------------------------------------------*
Private Function ParserEau(oNoeudDoc As MSXML2.IXMLDOMNode) As listeBouteille

    Dim oNoeudEau As MSXML2.IXMLDOMNode
    Dim iIcBouteille As Integer
    
    ' Nbre de bouteilles d'eau
    iIcBouteille = 0
    ' Pour chaque Noeud de "Eau"
    For Each oNoeudEau In oNoeudDoc.ChildNodes
        ' Noeud "Bouteille"
        If oNoeudEau.BaseName = "bouteille" Then
            ' Incrémenter le compteur des bouteilles
            ParserEau.NbBouteilles = iIcBouteille + 1
            ' Redimmensionner le tableau des bouteilles
            ReDim Preserve ParserEau.Bouteille(iIcBouteille) As Bouteille
            ' Renseigner les infos de la bouteille
            ParserEau.Bouteille(iIcBouteille) = ParserBouteille(oNoeudEau)
            ' Incrémenter l'indice des bouteilles
            iIcBouteille = iIcBouteille + 1
        End If
    Next oNoeudEau
    
End Function

' *---------------------------------------------------------------------------------------------------*
' * Parcourir tous les élements du noeud "Bouteille"                                                  *
' *---------------------------------------------------------------------------------------------------*
Private Function ParserBouteille(oNoeudEau As MSXML2.IXMLDOMNode) As Bouteille

    Dim oNoeudBouteille As MSXML2.IXMLDOMNode
    Dim iCompo As Integer

    For Each oNoeudBouteille In oNoeudEau.ChildNodes
        Select Case oNoeudBouteille.BaseName
            Case "marque":
                ParserBouteille.marque = oNoeudBouteille.nodeTypedValue
            Case "composition":
                ParserBouteille.listeComposition = ParserComposition(oNoeudBouteille)
            Case "source":
                ParserBouteille.source = ParserSource(oNoeudBouteille)
            Case "code_barre":
                ParserBouteille.codeBarre = oNoeudBouteille.nodeTypedValue
            Case "contenance":
                ParserBouteille.contenance = oNoeudBouteille.nodeTypedValue
            Case "ph":
                ParserBouteille.ph = oNoeudBouteille.nodeTypedValue
        End Select
    Next oNoeudBouteille

End Function

' *---------------------------------------------------------------------------------------------------*
' * Parcourir tous les éléments du noeud "Composition"                                                *
' *---------------------------------------------------------------------------------------------------*
Private Function ParserComposition(oNoeudBouteille As MSXML2.IXMLDOMNode) As listeComposition

    Dim oNoeudComposition As MSXML2.IXMLDOMNode, iComposition As Integer

    iComposition = 0
    For Each oNoeudComposition In oNoeudBouteille.ChildNodes
        If oNoeudComposition.BaseName = "ion" Then
            ParserComposition.nbIons = iComposition + 1
            ReDim Preserve ParserComposition.composition(iComposition) As composition
            ParserComposition.composition(iComposition).nomIon = oNoeudComposition.nodeTypedValue
            ParserComposition.composition(iComposition).typeIon = oNoeudComposition.Attributes.Item(0).nodeTypedValue
            iComposition = iComposition + 1
        End If
    Next oNoeudComposition

End Function

' *---------------------------------------------------------------------------------------------------*
' * Parcourir tous les éléments du noeud "source"                                                     *
' *---------------------------------------------------------------------------------------------------*
Private Function ParserSource(oNoeudBouteille As MSXML2.IXMLDOMNode) As source

    Dim oNoeudSource As MSXML2.IXMLDOMNode

    For Each oNoeudSource In oNoeudBouteille.ChildNodes
        Select Case oNoeudSource.BaseName
            Case "ville":
                ParserSource.ville = oNoeudSource.nodeTypedValue
            Case "departement":
                ParserSource.departement = oNoeudSource.nodeTypedValue
        End Select
    Next oNoeudSource

End Function

