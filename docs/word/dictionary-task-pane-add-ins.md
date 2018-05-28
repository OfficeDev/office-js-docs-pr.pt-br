---
title: Criar um suplemento de painel de tarefas de dicion?rio
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 781e2d07c88e56cbb64a7e7c5671dbbbc1b00894
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="create-a-dictionary-task-pane-add-in"></a>Criar um suplemento de painel de tarefas de dicion?rio


Este artigo mostra um exemplo de um suplemento de painel de tarefas e o servi?o Web correspondente que fornece defini??es de dicion?rio ou sin?nimos de dicion?rio de sin?nimos para a sele??o do usu?rio atual em um documento do Word 2013. 

Um Suplemento do Office de dicion?rio baseia-se no suplemento de painel de tarefas padr?o, com recursos adicionais para dar suporte a consultas e exibir defini??es de um servi?o Web XML de dicion?rio em locais adicionais na interface do usu?rio do aplicativo do Office. 

Em um suplemento de painel de tarefas de dicion?rio t?pico, um usu?rio seleciona uma palavra ou frase no documento e a l?gica de JavaScript por tr?s do suplemento passa essa sele??o ao servi?o Web XML do provedor do dicion?rio. A p?gina Web do provedor do dicion?rio ent?o ? atualizada para mostrar as defini??es para a sele??o ao usu?rio. O componente do servi?o Web XML retorna at? tr?s defini??es no formato definido pelo esquema OfficeDefinitions XML, que s?o exibidas para o usu?rio em outros locais na interface do usu?rio do aplicativo host do Office. A Figura 1 mostra a experi?ncia de sele??o e exibi??o para um suplemento de dicion?rio com a marca do Bing que est? em execu??o no Word 2013.

*Figura 1. Suplemento de dicion?rio exibindo defini??es para a palavra selecionada*

![Um aplicativo de dicion?rio exibindo uma defini??o](../images/dictionary-agave-01.jpg)

Voc? determina se clicar no link **Ver Mais** na interface do usu?rio HTML do suplemento de dicion?rio exibe mais informa??es no painel de tarefas ou abre uma janela separada do navegador para a p?gina da Web completa para a palavra ou frase selecionada. A Figura 2 mostra o comando do menu de contexto **Definir** que habilita os usu?rios a iniciar rapidamente os dicion?rios instalados. As Figuras 3 a 5 mostram os locais na interface do usu?rio do Office em que os servi?os de dicion?rio XML s?o usados para fornecer defini??es no Word 2013.

*Figura 2. Comando Definir no menu de contexto*

![Menu de contexto de Definir](../images/dictionary-agave-02.jpg)


*Figura 3. Defini??es nos pain?is Ortografia e Gram?tica*

![Defini??es nos pain?is Ortografia e Gram?tica](../images/dictionary-agave-03.jpg)


*Figura 4. Defini??es no painel Dicion?rio de Sin?nimos*

![Defini??es no painel Dicion?rio de Sin?nimos](../images/dictionary-agave-04.jpg)


*Figura 5. Defini??es no Modo de Leitura*

![Defini??es em modo de leitura](../images/dictionary-agave-05.jpg)

Para criar um suplemento de painel de tarefas que forne?a uma pesquisa de dicion?rio, crie dois componentes principais: 


- Um servi?o Web XML que pesquisa defini??es de um servi?o de dicion?rio e, em seguida, retorna os valores em um formato XML que pode ser consumido e exibido pelo suplemento de dicion?rio.
    
- Um suplemento de painel de tarefas que envia a sele??o atual do usu?rio ao servi?o Web de dicion?rio, exibe defini??es e, opcionalmente, pode inserir esses valores no documento.
    
As se??es a seguir fornecem exemplos de como criar esses componentes.

## <a name="creating-a-dictionary-xml-web-service"></a>Criar um servi?o Web XML de dicion?rio


O servi?o Web XML deve retornar consultas ao servi?o Web como XML que estejam de acordo com o esquema XML OfficeDefinitions. As duas se??es a seguir descrevem o esquema XML OfficeDefinitions e fornecem um exemplo de como escrever c?digo para um servi?o Web XML que retorna consultas nesse formato XML.


### <a name="officedefinitions-xml-schema"></a>Esquema XML OfficeDefinitions

O c?digo a seguir mostra o XSD para o esquema XML OfficeDefinitions.


```XML
<?xml version="1.0" encoding="utf-8"?>
<xs:schema
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  targetNamespace="http://schemas.microsoft.com/NLG/2011/OfficeDefinitions"
  xmlns="http://schemas.microsoft.com/NLG/2011/OfficeDefinitions">
  <xs:element name="Result">
    <xs:complexType>
      <xs:sequence>
        <xs:element name="SeeMoreURL" type="xs:anyURI"/>
        <xs:element name="Definitions" type="DefinitionListType"/>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
  <xs:complexType name="DefinitionListType">
    <xs:sequence>
      <xs:element name="Definition" maxOccurs="3">
        <xs:simpleType>
          <xs:restriction base="xs:normalizedString">
            <xs:maxLength value="400"/>
          </xs:restriction>
        </xs:simpleType>
      </xs:element>
    </xs:sequence>
  </xs:complexType>
</xs:schema>
```

O XML retornado que est? de acordo com o esquema OfficeDefinitions consiste em um elemento raiz **Result** que cont?m um elemento **Definitions** com zero a tr?s elementos filho **Definition**, cada um dos quais cont?m defini??es com no m?ximo 400 caracteres. Al?m disso, a URL da p?gina completa no site do dicion?rio deve ser fornecida com o elemento **SeeMoreURL**. O exemplo a seguir mostra a estrutura do XML retornado que est? em conformidade com o esquema OfficeDefinitions.

```XML
<?xml version="1.0" encoding="utf-8"?>
<Result xmlns="http://schemas.microsoft.com/NLG/2011/OfficeDefinitions">
  <SeeMoreURL xmlns="">www.bing.com/dictionary/search?q=example</SeeMoreURL>
  <Definitions xmlns="">
    <Definition>Definition1</Definition>
    <Definition>Definition2</Definition>
    <Definition>Definition3</Definition>
  </Definitions>
 </Result>

```


### <a name="sample-dictionary-xml-web-service"></a>Servi?o Web XML de dicion?rio de exemplo

O c?digo C# a seguir fornece um exemplo simples de como escrever c?digo para um servi?o Web XML que retorna o resultado de uma consulta ao dicion?rio no formato XML OfficeDefinitions.


```cs
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.Xml;
using System.Text;
using System.IO;
using System.Net;

/// <summary>
/// Summary description for _Default
/// </summary>
[WebService(Namespace = "http://tempuri.org/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
// To allow this web service to be called from script, using ASP.NET AJAX, uncomment the following line. 
// [System.Web.Script.Services.ScriptService]
public class WebService : System.Web.Services.WebService {

    public WebService () {

        // Uncomment the following line if using designed components 
        // InitializeComponent(); 
    }

    // You can replace this method entirely with your own method that gets definitions
    // from your data source, and then formats it into the OfficeDefinitions XML format. 
    // If you need a reference for constructing the returned XML, you can use this example as a basis.
    [WebMethod]
    public XmlDocument Define(string word)
    {

        StringBuilder sb = new StringBuilder();
        XmlWriter writer = XmlWriter.Create(sb);
        {
            writer.WriteStartDocument();
            
                writer.WriteStartElement("Result", "http://schemas.microsoft.com/NLG/2011/OfficeDefinitions");

            // See More URL should be changed to the dictionary publisher's page for that word on their website.
                    writer.WriteElementString("SeeMoreURL", "http://www.bing.com/search?q=" + word);

                    writer.WriteStartElement("Definitions");
            
                        writer.WriteElementString("Definition", "Definition 1 of " + word);
                        writer.WriteElementString("Definition", "Definition 2 of " + word);
                        writer.WriteElementString("Definition", "Definition 3 of " + word);
                   
                    writer.WriteEndElement();


                writer.WriteEndElement();
            
            writer.WriteEndDocument();
        }
        writer.Close();

        XmlDocument doc = new XmlDocument();
        doc.LoadXml(sb.ToString());

        return doc;
    }
}
```


## <a name="creating-the-components-of-a-dictionary-add-in"></a>Criar os componentes de um suplemento de dicion?rio


Um suplemento de dicion?rio consiste em tr?s arquivos de componentes principais:


- Um arquivo de manifesto XML que descreve o suplemento.
    
- Um arquivo HTML que fornece a interface do usu?rio do suplemento.
    
- Um arquivo JavaScript que fornece a l?gica para obter a sele??o do usu?rio do documento, envia a sele??o como uma consulta ao servi?o Web e exibe os resultados retornados na interface do usu?rio do suplemento.
    

### <a name="creating-a-dictionary-add-ins-manifest-file"></a>Criar um arquivo de manifesto de um suplemento de dicion?rio

A seguir h? um arquivo de manifesto de exemplo para um suplemento de dicion?rio.


```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <Id>7164e750-dc86-49c0-b548-1bac57abdc7c</Id>
  <Version>15.0</Version>
  <ProviderName>Microsoft Office Demo Dictionary</ProviderName>
  <DefaultLocale>en-us</DefaultLocale>
  <!--DisplayName is the name that will appear in the user's list of applications.-->
  <DisplayName DefaultValue="Microsoft Office Demo Dictionary" />
  <!--Description is a 2-3 sentence description of this dictionary. -->
  <Description DefaultValue="The Microsoft Office Demo Dictionary is an example built to demonstrate how a publisher could create a dictionary that integrates with Office. It does not return real definitions." />
  <!--IconUrl is the URI for the icon that will appear in the user's list of applications.-->
  <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg" />
  <!--Capabilities specifies the kind of host application your dictionary add-in will support. You shouldn't have to modify this area.-->
  <Capabilities>
    <Capability Name="Workbook"/>
    <Capability Name="Document"/>
    <Capability Name="Project"/>
  </Capabilities>
  <DefaultSettings>
    <!--SourceLocation is the URL for your dictionary-->
    <SourceLocation DefaultValue="http://christophernlg/ExampleDictionary/DictionaryHome.html" />
  </DefaultSettings>
  <!--Permissions is the set of permissions a user will have to give your dictionary. If you need write access, such as to allow a user to replace the highlighted word with a synonym, use ReadWriteDocument. -->
  <Permissions>ReadDocument</Permissions>
  <Dictionary>
    <!--TargetDialects is the set of regional languages your dictionary contains. For example, if your dictionary applies to Spanish (Mexico) and Spanish (Peru), but not Spanish (Spain), you can specify that here. Do not put more than one language (for example, Spanish and English) here. Publish separate languages as separate dictionaries. -->
    <TargetDialects>
      <TargetDialect>EN-AU</TargetDialect>
      <TargetDialect>EN-BZ</TargetDialect>
      <TargetDialect>EN-CA</TargetDialect>
      <TargetDialect>EN-029</TargetDialect>
      <TargetDialect>EN-HK</TargetDialect>
      <TargetDialect>EN-IN</TargetDialect>
      <TargetDialect>EN-ID</TargetDialect>
      <TargetDialect>EN-IE</TargetDialect>
      <TargetDialect>EN-JM</TargetDialect>
      <TargetDialect>EN-MY</TargetDialect>
      <TargetDialect>EN-NZ</TargetDialect>
      <TargetDialect>EN-PH</TargetDialect>
      <TargetDialect>EN-SG</TargetDialect>
      <TargetDialect>EN-ZA</TargetDialect>
      <TargetDialect>EN-TT</TargetDialect>
      <TargetDialect>EN-GB</TargetDialect>
      <TargetDialect>EN-US</TargetDialect>
      <TargetDialect>EN-ZW</TargetDialect>
    </TargetDialects>
    <!--QueryUri is the address of this dictionary's XML web service (which is used to put definitions in additional contexts, such as the spelling checker.)-->
    <QueryUri DefaultValue="http://christophernlg/ExampleDictionary/WebService.asmx/Define?word="/>
    <!--Citation Text, Dictionary Name, and Dictionary Home Page will be combined to form the citation line (for example, this would produce "Examples by: Microsoft", where "Microsoft" is a hyperlink to http://www.microsoft.com).-->
    <CitationText DefaultValue="Examples by: " />
    <DictionaryName DefaultValue="Microsoft" />
    <DictionaryHomePage DefaultValue="http://www.microsoft.com" />
  </Dictionary>
</OfficeApp>
```

O elemento **Dictionary** e seus elementos filho que s?o espec?ficos para a cria??o do arquivo de manifesto de um suplemento de dicion?rio s?o descritos nas se??es a seguir. Para obter informa??es sobre os outros elementos no arquivo de manifesto, confira [Manifesto XML de suplementos do Office](../develop/add-in-manifests.md).


### <a name="dictionary-element"></a>Elemento Dictionary


Especifica configura??es para suplementos de dicion?rio.

 **Elemento pai**

 `<OfficeApp>`

 **Elementos filho**

 `<TargetDialects>`,  `<QueryUri>`,  `<CitationText>`,  `<DictionaryName>`,  `<DictionaryHomePage>`

 **Coment?rios**

O elemento **Dictionary** e seus elementos filho s?o adicionados ao manifesto de um suplemento de painel de tarefas ao criar um suplemento de dicion?rio.


#### <a name="targetdialects-element"></a>Elemento TargetDialects


Especifica os idiomas regionais aos quais o dicion?rio oferece suporte. Necess?rio para suplementos de dicion?rio.

 **Elemento pai**

 `<Dictionary>`

 **Elemento filho**

 `<TargetDialect>`

 **Coment?rios**

O elemento **TargetDialects** e os elementos filho dele especificam o conjunto de idiomas regionais que o dicion?rio cont?m. Por exemplo, se o dicion?rio se aplica a Espanhol (M?xico) e Espanhol (Peru), mas n?o a Espanhol (Espanha), ? poss?vel especificar isso nesse elemento. N?o especifique mais de um idioma (por exemplo, espanhol e ingl?s) nesse manifesto. Publique idiomas separados como dicion?rios separados.

 **Exemplo**

```XML
<TargetDialects>
  <TargetDialect>EN-AU</TargetDialect>
  <TargetDialect>EN-BZ</TargetDialect>
  <TargetDialect>EN-CA</TargetDialect>
  <TargetDialect>EN-029</TargetDialect>
  <TargetDialect>EN-HK</TargetDialect>
  <TargetDialect>EN-IN</TargetDialect>
  <TargetDialect>EN-ID</TargetDialect>
  <TargetDialect>EN-IE</TargetDialect>
  <TargetDialect>EN-JM</TargetDialect>
  <TargetDialect>EN-MY</TargetDialect>
  <TargetDialect>EN-NZ</TargetDialect>
  <TargetDialect>EN-PH</TargetDialect>
  <TargetDialect>EN-SG</TargetDialect>
  <TargetDialect>EN-ZA</TargetDialect>
  <TargetDialect>EN-TT</TargetDialect>
  <TargetDialect>EN-GB</TargetDialect>
  <TargetDialect>EN-US</TargetDialect>
  <TargetDialect>EN-ZW</TargetDialect>
</TargetDialects>
```


#### <a name="targetdialect-element"></a>Elemento TargetDialect


Especifica um idioma regional ao qual o dicion?rio oferece suporte. Necess?rio para suplementos de dicion?rio.

 **Elemento pai**

 `<TargetDialects>`

 **Coment?rios**

Especifique o valor para um idioma regional no formato de tag de `language` RFC1766, como PT-BR.

 **Exemplo**


```XML
<TargetDialect>EN-US</TargetDialect>
```


#### <a name="queryuri-element"></a>Elemento QueryUri


Especifica o ponto de extremidade do servi?o de consulta de dicion?rio. Necess?rio para suplementos de dicion?rio.

 **Elemento pai**

 `<Dictionary>`

 **Coment?rios**

Esse ? o URI do servi?o Web XML para o provedor do dicion?rio. A consulta com escape correto ser? anexada a esse URI. 

 **Exemplo**


```XML
<QueryUri DefaultValue="http://msranlc-lingo1/proof.aspx?q="/>
```


#### <a name="citationtext-element"></a>Elemento CitationText


Especifica o texto a ser usado em cita??es. Necess?rio para suplementos de dicion?rio.

 **Elemento pai**

 `<Dictionary>`

 **Coment?rios**

Esse elemento especifica o in?cio do texto de cita??o que ser? exibido em uma linha abaixo do conte?do que ? retornado do servi?o Web (por exemplo, "Resultados do:" ou "Da plataforma:").

Para esse elemento, voc? pode especificar valores para localidades adicionais usando o elemento **Override**. Por exemplo, se um usu?rio est? executando a SKU do portugu?s brasileiro do Office, mas usando um dicion?rio de ingl?s, isso permite que a linha de cita??o seja "Resultados por: Bing"em vez de "Results by: Bing". Para saber mais sobre como especificar valores para localidades adicionais, confira a se??o "Fornecer configura??es para localidades diferentes" em [Manifesto XML de suplementos do Office](../develop/add-in-manifests.md).

 **Exemplo**


```XML
<CitationText DefaultValue="Results by: " />
```


#### <a name="dictionaryname-element"></a>Elemento DictionaryName


Especifica o nome deste dicion?rio. Necess?rio para suplementos de dicion?rio.

 **Elemento pai**

 `<Dictionary>`

 **Coment?rios**

Esse elemento especifica o texto do link no texto de cita??o. O texto de cita??o ? exibido em uma linha abaixo do conte?do que ? retornado do servi?o Web.

Para esse elemento, voc? pode especificar valores para localidades adicionais.

 **Exemplo**

```XML
<DictionaryName DefaultValue="Bing Dictionary" />
```


#### <a name="dictionaryhomepage-element"></a>Elemento DictionaryHomePage


Especifica a URL da p?gina inicial do dicion?rio. Necess?rio para suplementos de dicion?rio.

 **Elemento pai**

 `<Dictionary>`

 **Coment?rios**

Esse elemento especifica a URL do link no texto de cita??o. O texto de cita??o ? exibido em uma linha abaixo do conte?do que ? retornado do servi?o Web.

Para esse elemento, voc? pode especificar valores para localidades adicionais.

 **Exemplo**


```XML
<DictionaryHomePage DefaultValue="http://www.bing.com" />
```


### <a name="creating-a-dictionary-add-ins-html-user-interface"></a>Criar a interface do usu?rio HTML de um suplemento de dicion?rio

Os dois exemplos a seguir mostram os arquivos HTML e CSS para a interface do usu?rio do suplemento de Dicion?rio de Demonstra??o. Para ver como a interface do usu?rio ? exibida no suplemento de painel de tarefas, confira a Figura 6 ap?s o c?digo. Para ver como a implementa??o do JavaScript no arquivo Dictionary.js fornece l?gica de programa??o para essa interface do usu?rio HTML, confira "Escrever a implementa??o de JavaScript" imediatamente ap?s esta se??o.

```HTML
<!DOCTYPE html>
<html>

<head>
<meta http-equiv="X-UA-Compatible" content="IE=Edge"/>

<!--The title will not be shown but is supplied to ensure valid HTML.-->
<title>Example Dictionary</title>

<!--Required library includes.-->
<script type="text/javascript" src="http://ajax.microsoft.com/ajax/4.0/1/MicrosoftAjax.js"></script>
<script type="text/javascript" src="office.js"></script>

<!--Optional library includes.-->
<script type="text/javascript" src="http://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.5.1.js"></script>

<!--App-specific CSS and JS.-->
<link rel="Stylesheet" type="text/css" href="style.css" />
<script type="text/ecmascript" src="dictionary.js"></script>
</head>

<body>
<div id="mainContainer">
    <div id="header">
        <span id="headword"></span>
        <span id="pronunciation">(<a id="pronunciationLink">Pronounce</a>)</span>
    </div>
    <ol id="definitions">
    </ol>
    <div id="SeeMore">
    <a id="SeeMoreLink">See More...</a>
    </div>
</div>
</body>

</html>
```

O exemplo a seguir mostra o conte?do de Style.css.

```CSS
#mainContainer
{
    font-family: Segoe UI;
    font-size: 11pt;
}

#headword
{
    font-family: Segoe UI Semibold;
    color: #262626;
}

#pronunciation
{
    margin-left: 2px;
    margin-right: 2px;
}

#definitions
{
    font-size: 8.5pt;
}
a
{
    font-size: 8pt;
    color: #336699;
    text-decoration: none;
}
a:visited
{
    color: #993366;
}
a:hover, a:active
{  
    text-decoration: underline;
}
```

*Figura 6. Demonstra??o da interface de usu?rio do dicion?rio*

![Demonstra??o da interface de usu?rio do dicion?rio](../images/dictionary-agave-06.jpg)


### <a name="writing-the-javascript-implementation"></a>Escrever a implementa??o de JavaScript


O exemplo a seguir mostra a implementa??o de JavaScript no arquivo Dictionary.js que ? chamada da p?gina HTML do suplemento para fornecer a l?gica de programa??o ao suplemento de Dicion?rio de Demonstra??o. Esse script reutiliza o servi?o Web XML descrito anteriormente. Quando colocado no mesmo diret?rio que o servi?o Web de exemplo, o script obter? defini??es desse servi?o. Para us?-lo com um servi?o Web XML p?blico em conformidade com OfficeDefinitions, modifique a vari?vel `xmlServiceURL` no in?cio do arquivo e substitua a chave API do Bing para pron?ncias com um script registrado corretamente.

Os membros prim?rios da API JavaScript para Office (Office.js) que s?o chamados por essa implementa??o s?o os seguintes:


- O evento [initialize](https://dev.office.com/reference/add-ins/shared/office.initialize) do objeto **Office**, que ? gerado quando o contexto do suplemento ? inicializado e fornece acesso a uma inst?ncia de objeto [Document](https://dev.office.com/reference/add-ins/shared/document) que representa o documento com o qual o suplemento est? interagindo.
    
- O m?todo [addHandlerAsync](https://dev.office.com/reference/add-ins/shared/document.addhandlerasync) do objeto **Document**, que ? chamado na fun??o **initialize** para adicionar um manipulador de eventos ao evento [SelectionChanged](https://dev.office.com/reference/add-ins/shared/document.selectionchanged.event) do documento para escutar altera??es de sele??o de usu?rio.
    
- O m?todo [getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) do objeto **Document**, que ? chamado na fun??o `tryUpdatingSelectedWord()` quando o manipulador de eventos **SelectionChanged** ? gerado para obter a palavra ou frase que o usu?rio selecionou, fazer a coer??o dela para texto sem formata??o e executar a fun??o `selectedTextCallback` de retorno de chamada ass?ncrono.
    
- Quando a fun??o de retorno de chamada ass?ncrono `selectTextCallback` que ? passada como o argumento _callback_ do m?todo **getSelectedDataAsync** ? executada, obt?m o valor do texto selecionado quando o retorno de chamada retorna. Ela obt?m o valor do argumento _selectedText_ do retorno de chamada (que ? do tipo [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult)) usando a propriedade [value](https://dev.office.com/reference/add-ins/shared/asyncresult.status) do objeto **AsyncResult** retornado.
    
- O restante do c?digo na fun??o `selectedTextCallback` consulta o servi?o Web XML para obter defini??es. Tamb?m chama as APIs do Microsoft Translator para fornecer a URL de um arquivo .wav que tem a pron?ncia da palavra selecionada.
    
- O c?digo restante em Dictionary.js exibe a lista de defini??es e o link de pron?ncia na interface do usu?rio HTML do suplemento.
    



```javascript
// The document the dictionary add-in is interacting with.
var _doc; 
// The last looked-up word, which is also the currently displayed word.
var lastLookup; 
// For demo purposes only!! Get an AppID if you intend to use the Pronunciation service for your feature.
var appID="3D8D4E1888B88B975484F0CA25CDD24AAC457ED8"; 

// The base URL for the OfficeDefinitions-conforming XML web service to query for definitions.
var xmlServiceUrl = "WebService.asmx/Define?Word="; 

// Initialize the add-in. 
// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    // Store a reference to the current document.
    _doc = Office.context.document; 
    // Check whether text is already selected.
    tryUpdatingSelectedWord(); 
    _doc.addHandlerAsync("documentSelectionChanged", tryUpdatingSelectedWord); //Add a handler to refresh when the user changes selection.
    });
}

// Executes when event is raised on user's selection changes, and at initialization time. 
// Gets the current selection and passes that to asynchronous callback method.
function tryUpdatingSelectedWord() {
    _doc.getSelectedDataAsync(Office.CoercionType.Text, selectedTextCallback); 
}

// Async callback that executes when the add-in gets the user's selection.
// Determines whether anything should be done. If so, it makes requests that will be passed to various functions.
function selectedTextCallback(selectedText) {
    selectedText = $.trim(selectedText.value);
    // Be sure user has selected text. The SelectionChanged event is raised every time the user moves the cursor, even if no selection.
    if (selectedText != "") { 
        // Check whether user selected the same word the pane is currently displaying to avoid unnecessary web calls.
        if (selectedText != lastLookup) { 
            // Update the lastLookup variable.
            lastLookup = selectedText; 
            // Set the "headword" span to the word you looked up.
            $("#headword").text(selectedText); 
            // AJAX request to get definitions for the selected word; pass that to refreshDefinitions.
            $.ajax(xmlServiceUrl + selectedText, { dataType: 'xml', success: refreshDefinitions, error: errorHandler }); 
            // AJAX request to the Microsoft Translator APIs. Gets the URL of a WAV file with pronunciation, which is passed to refreshPronunciation. See http://www.microsofttranslator.com/dev for details.
            $.ajax("http://api.microsofttranslator.com/V2/Ajax.svc/Speak?oncomplete=refreshPronunciation&amp;appId=" + appID + "&amp;text=" + selectedText + "&amp;language=en-us", { dataType: 'script', success: null, error: errorHandler }); 
        }
    }
}

// This function is called when the add-in gets back the definitions target word.
// It removes the old definitions and replaces them with the definitions for the current word.
// It also sets the "See More" link.
function refreshDefinitions(data, textStatus, jqXHR) {
    $(".definition").remove();
    // Make a new list item for each returned definition that was returned, set the CSS class, and append it to the definitions div.
    $(data).find("Definition").each(function () {
        $(document.createElement("li")).text($(this).text()).addClass("definition").appendTo($("#definitions"));
    });
    $("#SeeMoreLink").attr("href", $(data).find("SeeMoreURL").text()); //Change the "See More" link to direct to the correct URL.
}

// This function is called when the add-in gets back the link to the pronunciation
// to set the "Pronounce" link to the URL of the .WAV file.
function refreshPronunciation(data) {
    $("#pronunciationLink").attr("href", data);
}

// Basic error handler that writes to a div with id='message'.
function errorHandler(jqXHR, textStatus, errorThrown) {
    document.getElementById('message').innerText += errorThrown;
}

```

