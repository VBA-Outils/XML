# XML
<h1>Licence</h1>
<p>Ce projet est distribué sous licence MIT. Consultez le fichier LICENSE pour plus de détails.</p>
<h1>Prérequis</h1>
<p>Environnement de développement : Microsoft Visual Basic for Applications (VBA)</p>
<h1>Introduction</h1>
<p>Ce projet "XML" fournit des exemples d'utilisation des API Microsoft XML3.0 qui permettent de gérer des flux XML (parser, créer / supprimer des noeuds, indenter, etc). Il n'est pas exhaustif quant aux possibilités offertes par ces API.</p>
<p>Il est composé de 2 exemples :
<ul>
  <li>Chargement d'un flux XML, puis appel de la fonction de Parse, et enfin extraction des informations.</li>
  <li>Ouvrir un fichier MXL existant, lui ajouter des noeuds (afin de l'insérer dans une structure SOAPUI), puis indentation du flux final</li>
</ul></p>
<h1>Fonctionnalités :</h1>
<h2>Parser un flux XML</h2>
<p>Un flux XML peut être parsé en VBA en utilisant la référence disponible dans VBA : Microsoft XML, V3.0.</p>
<p>Un objet XML est généré lors du chargement du flux MXL, il peut ensuite être parcouru afin d'extraire les informations présentes.</p>
<p>Le module "ParserXML.bas" permet d'extraire les données du flux "Eaux.XML" dans une structure VBA. Il est appelé par la procédure "ParserFichierXML" présente dans le module "ParserFlux.bas", cette dernière va ouvrir une boîte de dialogue de sélection d'un fichier puis appeler la fonction qui réalise le Parse.</p>
<h2>Créer un nouveau flux XML  partir d'un existant et l'indenter</h2>
<p>Le module "CreerFluxXML" charge le flux Eau.xml, puis crée un nouvel objet XML, lui ajoute les caractéristiques des flux SOAPUI puis insère dans le corps le flux chargé précédemment.</p>
<p>Il appelle ensuite la classe "IndenterXML" qui permet d'indenter un flux MXL grâce aux API MXXMLWriter et MXXMLReader.</p>
