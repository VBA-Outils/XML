Attribute VB_Name = "ParserFlux"
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

Option Explicit

' *---------------------------------------------------------------------------------------------------*
' * Parser un fichier XML                                                                             *
' *---------------------------------------------------------------------------------------------------*
Public Sub ParserFichierXML()

    Dim sNomFichier As String, listeBouteille As listeBouteille
    Dim oFichier As New ADODB
    
    ' Sélection du fichier à parser
    With oFichier
        .Repertoire = Environ("OneDrive") & "\Documents\"
        'If .RepertoireExiste = True Then .NomInitialFichier = .Repertoire
        .ExtensionFiltre = "*.xml"
        .LibelleFiltre = "Fichiers XML"
        .TitreBoiteDeDialogue = "Sélectionner un Flux XML"
        .SelectionnerFichier
        If .NomFichier = "" Then Exit Sub
        sNomFichier = .NomFichier
    End With

    ' Désactiver la réactualisation de l'écran (performances)
    Call InitialiserTraitement
    
    listeBouteille = ParserFluxXML(NomFichier:=sNomFichier)

    Call TerminerTraitement
    
End Sub

