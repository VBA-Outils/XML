# XML
<h1>Utiliser des flux XML en VBA (parser, créer, indenter)</h1>

<h2>Parser un flux XML</h2>
<p>Un flux XML peut être parsé en VBA en utilisant les références disponibles dans VBA : Microsoft XML, V3.0.</p>
<p>Un objet XML est généré lors du chargement du flux MXL, il peut ensuite être parcouru afin d'extraire les informations présentes.</p>
<p>Le module "ParserXML.bas" permet d'extraire les données du flux "Eaux.XML" dans une structure VBA. Il est appelé par la procédure "ParserFichierXML" présente dans le module "ParserFlux.bas", cette dernière va ouvrir une boîte de dialogue de sélection d'un fichier puis appeler la fonction qui réalise le Parse.</p>
