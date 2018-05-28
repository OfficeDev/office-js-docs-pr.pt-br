---
title: Localiza??o para Suplementos do Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: d7888859ca29a62541020b45b0b7a3638c41f4f2
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="localization-for-office-add-ins"></a>Localiza??o para Suplementos do Office

Voc? pode implementar qualquer esquema de localiza??o que seja apropriado para o seu Suplemento do Office. A API JavaScript e o esquema do manifesto da plataforma de Suplementos do Office oferecem algumas op??es. Voc? pode usar a API JavaScript para Office para determinar uma localidade e exibir cadeias de caracteres com base na localidade do aplicativo host ou para interpretar ou exibir dados com base na localidade dos dados. Voc? pode usar o manifesto para especificar informa??es descritivas e o local do arquivo do suplemento espec?fico da localidade. Como alternativa, voc? pode usar o script do Microsoft Ajax para dar suporte ? globaliza??o e localiza??o.

## <a name="use-the-javascript-api-to-determine-locale-specific-strings"></a>Usar a API JavaScript para determinar cadeias de caracteres espec?ficas da localidade

A API JavaScript para Office fornece duas propriedades que oferecem suporte ? exibi??o ou interpreta??o de valores consistentes com a localidade do aplicativo host e dos dados:

- [Context.displayLanguage][displayLanguage] especifica a localidade (ou idioma) da interface do usu?rio do aplicativo host. O exemplo a seguir verifica se o aplicativo host usa a localidade en-US ou fr-FR e exibe uma sauda??o espec?fica para a localidade.
    
    ```js
    function sayHelloWithDisplayLanguage() {
        var myLanguage = Office.context.displayLanguage;
        switch (myLanguage) {
            case 'en-US':
                write('Hello!');
                break;
            case 'fr-FR':
                write('Bonjour!');
                break;
        }
    }
    
    // Function that writes to a div with id='message' on the page.
    function write(message) {
        document.getElementById('message').innerText += message; 
    }
    ```

- [Context.contentLanguage][contentLanguage] especifica a localidade (ou o idioma) dos dados. Estendendo o ?ltimo exemplo de c?digo, em vez de verificar a propriedade [displayLanguage], atribua `myLanguage` ? propriedade [contentLanguage] e use o restante do mesmo c?digo para exibir uma sauda??o com base na localidade dos dados:
    
    ```js
    var myLanguage = Office.context.contentLanguage;
    ```

## <a name="control-localization-from-the-manifest"></a>Controlar a localiza??o do manifesto


Cada Suplemento do Office especifica um elemento [DefaultLocale] e uma localidade em seu manifesto. Por padr?o, a plataforma do Suplemento do Office e os aplicativos host do Office aplicam os valores dos elementos [Description], [DisplayName], [IconUrl], [HighResolutionIconUrl] e [SourceLocation] a todas as localidades. Como op??o, voc? pode dar suporte a valores espec?ficos para localidades espec?ficas, especificando um elemento-filho [Override]para cada localidade adicional, para qualquer um desses cinco elementos. O valor do elemento [DefaultLocale] e do atributo `Locale` do elemento [Override] ? especificado de acordo com o [RFC 3066], "Marcas para a Identifica??o dos Idiomas". A Tabela 1 descreve o suporte de localiza??o para esses elementos.

**Tabela 1. Suporte de localiza??o**


|**Elemento**|**Suporte de localiza??o**|
|:-----|:-----|
|[Descri??o]   |Os usu?rios de cada localidade especificada podem ver uma descri??o localizada do suplemento no AppSource (ou no cat?logo privado).<br/>Para os suplementos do Outlook, os usu?rios podem ver a descri??o no Centro de Administra??o do Exchange (EAC) ap?s a instala??o.|
|[DisplayName]   |Os usu?rios de cada localidade especificada podem ver uma descri??o localizada do suplemento no AppSource (ou no cat?logo privado).<br/>Para os suplementos do Outlook, os usu?rios podem ver o nome de exibi??o como um r?tulo para o bot?o de suplemento do Outlook e no EAC ap?s a instala??o.<br/>Para os suplementos do painel de tarefas e do conte?do, os usu?rios podem ver o nome de exibi??o na faixa de op??es ap?s a instala??o do suplemento.|
|[IconUrl]        |A imagem do ?cone ? opcional. Voc? pode usar a mesma t?cnica de substitui??o para especificar uma determinada imagem para uma cultura espec?fica. Se voc? usar e localizar um ?cone, os usu?rios em cada localidade que voc? especificar poder?o ver uma imagem de ?cone localizada para o suplemento.<br/>Para suplementos do Outlook, os usu?rios podem ver o ?cone no EAC depois de instalar o suplemento.<br/>Para os suplementos do painel de tarefas e do conte?do, os usu?rios podem ver o ?cone na faixa de op??es ap?s a instala??o do suplemento.|
|[HighResolutionIconUrl] **Importante:** este elemento s? fica dispon?vel ao usar a vers?o 1.1 do manifesto do suplemento.|A imagem do ?cone de alta resolu??o ? opcional, mas se ela for especificada, dever? ocorrer ap?s o elemento [IconUrl]. Quando [HighResolutionIconUrl] for especificado e o suplemento estiver instalado em um dispositivo que ofere?a suporte ? resolu??o dpi alto, o valor [HighResolutionIconUrl] ? usado em vez do valor para [IconUrl].<br/>Voc? pode usar a mesma t?cnica de substitui??o para especificar uma determinada imagem para uma cultura espec?fica. Se voc? usar e localizar um ?cone, os usu?rios em cada localidade que voc? especificar podem ver uma imagem de ?cone localizada para o suplemento.<br/>Para suplementos do Outlook, os usu?rios podem ver o ?cone no EAC depois de instalar o suplemento.<br/>Para os suplementos do painel de tarefas e do conte?do, os usu?rios podem ver o ?cone na faixa de op??es ap?s a instala??o do suplemento.|
|[Recursos] **Importante:** este elemento s? fica dispon?vel ao usar a vers?o 1.1 do manifesto do suplemento.   |Os usu?rios em cada localidade especificada podem ver recursos de cadeias de caracteres e de ?cones que voc? projetou especificamente para o suplemento dessa localidade. |
|[SourceLocation]   |Os usu?rios em cada localidade especificada podem ver uma p?gina da Web que voc? projetou especificamente para o suplemento dessa localidade. |


> **OBSERVA??O:** voc? s? pode localizar o nome de exibi??o e a descri??o das localidades que oferecem suporte ao Office. Veja [Identificadores de idioma e valores de OptionState Id no Office 2013](http://technet.microsoft.com/en-us/library/cc179219.aspx) para obter uma lista de idiomas e localidades para a vers?o atual do Office.


### <a name="examples"></a>Exemplos

Por exemplo, um Suplemento do Office pode especificar o [DefaultLocale] como `en-us`. Para o elemento [DisplayName], o suplemento pode especificar um elemento filho [Override] para a localidade `fr-fr`, como mostrado abaixo. 


```xml
<DefaultLocale>en-us</DefaultLocale>
...
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

> **OBSERVA??O:** se for preciso localizar para mais de uma ?rea dentro de uma fam?lia de idiomas, como `de-de` e `de-at`, recomendamos que voc? use elementos `Override` separados para cada ?rea. Usar apenas o nome do idioma sozinho, nesse caso, `de`, n?o tem suporte em todas as combina??es de plataformas e aplicativos de host do Office.

Isso significa que o suplemento pressup?e a localidade `en-us` como padr?o. Os usu?rios veem o nome de exibi??o em ingl?s "Video player" para todas as localidades, a menos que a localidade do computador cliente fosse `fr-fr`, nesse caso os usu?rios veria o nome de exibi??o em franc?s "Lecteur vid?o".

> **Observa??o:** voc? s? pode especificar uma ?nica substitui??o por idioma, inclusive para a localidade padr?o. Por exemplo, se sua localidade padr?o ? `en-us`, n?o ? poss?vel especificar tamb?m uma substitui??o para `en-us`. 

O exemplo a seguir se aplica a uma substitui??o de localidade para o elemento [Description]. Primeiro especifica a localidade padr?o `en-us` e uma descri??o em ingl?s e, em seguida, especifica uma pol?tica de [Override] com uma descri??o francesa para a localidade `fr-fr`:

```xml
<DefaultLocale>en-us</DefaultLocale>
...
<Description DefaultValue=
   "Watch YouTube videos referenced in the emails you receive 
   without leaving your email client.">
   <Override Locale="fr-fr" Value=
   "Visualisez les vidéos YouTube référencées dans vos courriers 
   électronique directement depuis Outlook et Outlook Web App."/>
</Description>
```

Isso significa que o suplemento pressup?e a localidade `en-us` como padr?o. Os usu?rios veriam a descri??o em ingl?s no atributo `DefaultValue` para todas as localidades, a menos que a localidade do computador cliente fosse `fr-fr`, nesse caso, eles veriam a descri??o em franc?s.

No exemplo a seguir, o suplemento especifica uma imagem separada mais apropriada para a localidade e a cultura `fr-fr`. Os usu?rios ver?o a imagem DefaultLogo.png por padr?o, exceto quando a localidade do computador cliente for `fr-fr`. Nesse caso, os usu?rios veriam a imagem FrenchLogo.png. 


```xml
<!-- Replace "domain" with a real web server name and path. -->
<IconUrl DefaultValue="https://<domain>/DefaultLogo.png"/>
<Override Locale="fr-fr" Value="https://<domain>/FrenchLogo.png"/>
```

O exemplo a seguir mostra como localizar um recurso na se??o `Resources`. Ele aplica um substituto local para uma imagem que ? mais apropriada para a cultura `ja-jp`.

```xml
<Resources>
      <bt:Images>
        <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
        </bt:Image>
 ...
```


Para o elemento [SourceLocation], o suporte a localidades adicionais significa fornecer um arquivo HTML de origem separado para cada um dos locais especificados. Os usu?rios de cada localidade que voc? especificar poder?o ver uma p?gina da Web personalizada que foi projetada para eles.

Para suplementos do Outlook, o elemento [SourceLocation] tamb?m atribui o fator forma, o que permite que voc? forne?a um arquivo HTML de origem localizado e distinto para cada fator de foram correspondente. Voc? pode especificar um ou mais elementos filho [Override] em cada configura??o aplic?vel ([DesktopSettings], [TabletSettings] ou [PhoneSettings]). O exemplo a seguir mostra os elementos de configura??es para fatores de forma de desktop, tablet e smartphone, cada um com um arquivo HTML para a localidade padr?o e outro para a localidade francesa.


```xml
<DesktopSettings>
   <SourceLocation DefaultValue="https://contoso.com/Desktop.html">
      <Override Locale="fr-fr" Value="https://contoso.com/fr/Desktop.html" />
   </SourceLocation>
   <RequestedHeight>250</RequestedHeight>
</DesktopSettings>
<TabletSettings>
   <SourceLocation DefaultValue="https://contoso.com/Tablet.html">
      <Override Locale="fr-fr" Value="https://contoso.com/fr/Tablet.html" />
   </SourceLocation>
   <RequestedHeight>200</RequestedHeight>
</TabletSettings>
<PhoneSettings>
   <SourceLocation DefaultValue="https://contoso.com/Mobile.html">
      <Override Locale="fr-fr" Value="https://contoso.com/fr/Mobile.html" />
   </SourceLocation>
</PhoneSettings>
```

## <a name="match-datetime-format-with-client-locale"></a>Fazer a correspond?ncia entre o formato de data/hora e a localidade do cliente

Voc? pode obter a localidade da interface do usu?rio do aplicativo host usando a propriedade [displayLanguage]. Em seguida, pode exibir valores de data e hora em um formato consistente com a localidade atual do aplicativo host. Uma maneira de fazer isso ? preparar um arquivo de recurso que especifica o formato de exibi??o de data/hora a ser usado em cada localidade com suporte do seu Suplemento do Office. Na execu??o, seu suplemento pode usar o arquivo de recurso e fazer a correspond?ncia entre o formato de data/hora apropriado e a localidade obtida na propriedade [displayLanguage]

Voc? pode obter a localidade dos dados do aplicativo host usando a propriedade [contentLanguage]. Com base nesse valor, voc? pode, ent?o, interpretar ou exibir adequadamente as cadeias de caracteres de data/hora. Por exemplo, a localidade `jp-JP` expressa valores de data/hora como `yyyy/MM/dd`, e a localidade `fr-FR` como `dd/MM/yyyy`.


## <a name="use-ajax-for-globalization-and-localization"></a>Usar o Ajax para a globaliza??o e a localiza??o


Se voc? usar o Visual Studio para criar Suplementos do Office, o .NET Framework e Ajax fornecem maneiras de globalizar e localizar arquivos de script de cliente.

Voc? pode globalizar e utilizar as extens?es do tipo JavaScript de [Data](http://msdn.microsoft.com/library/caf98d32-2de2-4704-8198-692350343681.aspx) e [N?mero](http://msdn.microsoft.com/library/c216d3a1-12ae-47d1-bca1-c3666d04572f.aspx) e o objeto [Data](http://msdn.microsoft.com/library/ce2202bb-7ec9-4f5a-bf48-3a04feff283e.aspx) do JavaScript no c?digo do JavaScript para um suplemento do Office para exibir valores com base nas configura??es de localiza??o do navegador atual. Para saber mais, confira [Passo a passo: como globalizar uma data usando o script de cliente](http://msdn.microsoft.com/library/69b34e6d-d590-4d03-a763-b7ae54b47d74.aspx).

Voc? pode incluir cadeias de caracteres de recurso localizadas diretamente em arquivos de JavaScript aut?nomos para fornecer arquivos de script de cliente para diferentes locais, que s?o definidos no navegador ou fornecidos pelo usu?rio. Crie um arquivo de script separado para cada localidade com suporte. Em cada arquivo de script, inclua um objeto no formato JSON que contenha as cadeias de caracteres de recursos para essa localidade. Os valores localizados ser?o aplicados quando o script for executado no navegador. 


## <a name="example-build-a-localized-office-add-in"></a>Exemplo: Criar um Suplemento do Office localizado

Esta se??o fornece exemplos que mostram como localizar uma descri??o do Suplemento do Office, o nome de exibi??o e interface do usu?rio.

Para executar o c?digo de amostra fornecido, configure o Microsoft Office 2013 em seu computador para usar idiomas adicionais para que voc? possa testar seu suplemento, alternando o idioma usado para exibi??o em menus e em comandos para edi??o e revis?o de texto ou ambos.

Al?m disso, voc? precisar? criar um projeto de Suplemento do Office do Visual Studio 2015.

> **Observa??o:** para baixar o Visual Studio 2015, confira a [P?gina do Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs). Esta p?gina tamb?m tem um link para o Office Developer Tools.

### <a name="configure-office-2013-to-use-additional-languages-for-display-or-editing"></a>Configurar o Office 2013 para usar idiomas adicionais para exibi??o ou edi??o

Voc? pode usar um Pacote de idiomas do Office 2013 para instalar um idioma adicional. Para saber mais sobre os Pacotes de idioma e onde obt?-los, veja [Op??es de idioma do Office 2013](http://office.microsoft.com/en-us/language-packs/).

> **OBSERVA??O:** se voc? for assinante do MSDN, ? poss?vel que j? tenha os Pacotes de Idiomas do Office 2013. Para determinar se a sua assinatura oferece Pacotes de Idiomas do Office 2013 para download, v? para [P?gina Inicial de Assinaturas do MSDN](https://msdn.microsoft.com/subscriptions/manage/), insira Pacote de Idiomas do Office 2013 em **Downloads de Softwares**, escolha **Pesquisa** e selecione **Produtos dispon?veis com minha assinatura**. Em **Idioma**, marque a caixa de sele??o do Pacote de Idiomas que voc? deseja baixar e, em seguida, selecione **Ir**. 

Depois de instalar o Pacote de Idiomas, voc? pode configurar o Office 2013 para usar o idioma instalado para exibir na interface do usu?rio, para edi??o de conte?do do documento, ou ambos. O exemplo neste artigo usa uma instala??o do Office 2013 que tenha o Pacote de Idiomas do espanhol aplicado.

### <a name="create-an-office-add-in-project"></a>Criar um projeto de Suplemento do Office

1. No Visual Studio, escolha **Arquivo**  >  **Novo Projeto**.
    
2. Na caixa de di?logo **Novo Projeto**, em **Modelos**, expanda **Visual Basic** ou **Visual C#**, expanda **Office/SharePoint** e, em seguida, selecione **Suplementos do Office**.
    
3. Escolha **Suplemento do Office** e, em seguida, nomeie seu suplemento, por exemplo WorldReadyAddIn. Escolha **OK**.
    
4. Na caixa de di?logo **Criar Suplemento do Office**, selecione **Painel de tarefas** e selecione **Pr?ximo**. Na pr?xima p?gina, desmarque e marque as caixas de todos os aplicativos, exceto do Word. Selecione **Concluir** para criar o projeto.
    

### <a name="localize-the-text-used-in-your-add-in"></a>Localizar o texto usado no seu suplemento

O texto que voc? deseja localizar para outro idioma aparece em duas ?reas:

-  **Nome de exibi??o e descri??o do suplemento**. Isso ? controlado por entradas no arquivo do manifesto do suplemento.
    
-  **Interface do Usu?rio do Suplemento**. Voc? pode localizar as cadeias de caracteres que aparecem na interface do usu?rio do seu suplemento usando c?digos do JavaScript, por exemplo, usando um arquivo de recurso separado que contenha as cadeias de caracteres localizadas.
    
Para localizar o nome de exibi??o e a descri??o do suplemento:

1. Em **Gerenciador de Solu??es**, expanda **WorldReadyAddIn**, **WorldReadyAddInManifest** e, em seguida, selecione **WorldReadyAddIn.xml**.
    
2. No WorldReadyAddInManifest.xml, substitua os elementos [DisplayName] e [Description] com o seguinte bloqueio de c?digo:
    
    > **OBSERVA??O:** voc? pode substituir as cadeias de caracteres do idioma espanhol localizado usadas neste exemplo pelos elementos [DisplayName] e [Description] pelas cadeias de caracteres localizadas de qualquer outro idioma.

    ```xml
    <DisplayName DefaultValue="World Ready add-in">
      <Override Locale="es-es" Value="Aplicación de uso internacional"/>
    </DisplayName>
    <Description DefaultValue="An add-in for testing localization">
      <Override Locale="es-es" Value="Una aplicación para la prueba de la localización"/>
    </Description>
    ```

3. Quando voc? altera o idioma de exibi??o do Office 2013 do ingl?s para o espanhol, por exemplo, e executa o suplemento, o nome de exibi??o do suplemento e a descri??o s?o mostrados com texto localizado. 
    
Para definir a interface do usu?rio do suplemento:

1. No Visual Studio, no **Gerenciador de Solu??es**, selecione **Home.html**.
    
2. Substitua o HTML em Home.html pelo seguinte HTML.
    
    ```html
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title></title>
        <script src="../../Scripts/jquery-1.8.2.js" type="text/javascript"></script>
    
        <link href="../../Content/Office.css" rel="stylesheet" type="text/css" />
        <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
    
        <!-- To enable offline debugging using a local reference to Office.js, use:                        -->
        <!-- <script src="../../Scripts/Office/MicrosoftAjax.js" type="text/javascript"></script>          -->
        <!--    <script src="../../Scripts/Office/1.0/office.js" type="text/javascript"></script>          -->
    
        <link href="../App.css" rel="stylesheet" type="text/css" />
        <script src="../App.js" type="text/javascript"></script>
    
        <link href="Home.css" rel="stylesheet" type="text/css" />
        <script src="Home.js" type="text/javascript"></script> <body>
        <!-- Page content -->
        <div id="content-header">
            <div class="padding">
                <h1 id="greeting"></h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <div>
                    <p id="about"></p>
                </div>            
            </div>
        </div>
    </head>
    </html>
    ```

3. No Visual Studio, selecione **Arquivo**,  **Salvar Suplemento\Home\Home.html**.
    
A figura a seguir mostra o elemento do cabe?alho (h1) e o elemento do par?grafo (p) que exibir? o texto localizado quando seu suplemento de amostra for executado.

*Figura 1. A interface do usu?rio do suplemento*

![Interface de usu?rio do aplicativo com as se??es real?adas.](../images/office15-app-how-to-localize-fig03.png)

### <a name="add-the-resource-file-that-contains-the-localized-strings"></a>Adicionar o arquivo de recurso que cont?m as cadeias de caracteres localizadas

O arquivo de recurso do JavaScript cont?m as cadeias de caracteres usadas para a interface do usu?rio do suplemento. A interface do usu?rio do suplemento de amostra tem um elemento h1 que exibe uma sauda??o e um elemento p que apresenta o suplemento ao usu?rio. 

Para habilitar cadeias de caracteres para o cabe?alho e par?grafo, coloque as cadeias de caracteres em um arquivo de recurso separado. O arquivo de recurso cria um objeto do JavaScript que cont?m um objeto JSON (JavaScript Object Notation) separado para cada conjunto de cadeias de caracteres localizadas. O arquivo de recurso tamb?m fornece um m?todo para obter o objeto JSON apropriado de volta para uma determinada localidade. 

Para adicionar o arquivo de recurso ao projeto do suplemento:

1. No **Gerenciador de Solu??es** no Visual Studio, escolha a pasta **Suplemento** no projeto da Web para o suplemento de amostra e selecione **Adicionar**  >  **Arquivo JavaScript**.
    
2. Na caixa de di?logo **Especificar o nome do item**, insira UIStrings.js.
    
3. Adicione o c?digo a seguir ao arquivo UIStrings.js.

    ```js
    /* Store the locale-specific strings */
    
    var UIStrings = (function ()
    {
        "use strict";
    
        var UIStrings = {};
    
        // JSON object for English strings
        UIStrings.EN =
        {        
            "Greeting": "Welcome",
            "Introduction": "This is my localized add-in."        
        };
    
        // JSON object for Spanish strings
        UIStrings.ES =
        {        
            "Greeting": "Bienvenido",
            "Introduction": "Esta es mi aplicación localizada."
        };
    
        UIStrings.getLocaleStrings = function (locale)
        {
            var text;
            
            // Get the resource strings that match the language.
            switch (locale)
            {
                case 'en-US':
                    text = UIStrings.EN;
                    break;
                case 'es-ES':
                    text = UIStrings.ES;
                    break;
                default:
                    text = UIStrings.EN;
                    break;
            }
    
            return text;
        };
    
        return UIStrings;
    })();
    ```

O arquivo de recurso UIStrings.js cria o objeto, **UIStrings**, que cont?m as cadeias de caracteres localizadas para a interface do usu?rio do suplemento. 

### <a name="localize-the-text-used-for-the-add-in-ui"></a>Localizar o texto usado na interface do usu?rio do suplemento

Para usar o arquivo de recurso no seu suplemento, voc? precisar? adicionar a ele uma marca de script em Home.html. Quando Home.html for carregado, o UIStrings.js ser? executado e o objeto **UIStrings** que voc? utiliza para obter a cadeia de caracteres ficar? dispon?vel para seu c?digo. Adicione o seguinte HTML ? marca de cabe?alho do Home.html para tornar **UIStrings** dispon?vel para seu c?digo.

```html
<!-- Resource file for localized strings:                                                          -->
<script src="../UIStrings.js" type="text/javascript"></script>
```

Agora voc? pode usar o objeto **UIStrings** para definir as cadeias de caracteres da interface do usu?rio do seu suplemento.

Se voc? quiser alterar a localiza??o do seu suplemento com base no idioma usado para exibi??o nos menus e comandos no aplicativo host, use a propriedade **Office.context.displayLanguage** para obter a localidade desse idioma. Por exemplo, se o idioma do aplicativo host utilizar espanhol para exibir menus e comandos, a propriedade **Office.context.displayLanguage** retornar? o c?digo es-ES.

Se voc? quiser alterar a localiza??o do seu suplemento com base no idioma que est? sendo usado para editar o conte?do do documento, use a propriedade **Office.context.contentLanguage** para obter a localidade do idioma. Por exemplo, se o idioma do aplicativo host utilizar espanhol para editar o conte?do do documento, a propriedade **Office.context.contentLanguage** retornar? o c?digo es-ES.

Depois que voc? souber o idioma que o aplicativo host est? utilizando, ? poss?vel usar **UIStrings** para obter o conjunto de cadeias de caracteres localizadas correspondentes ao idioma do aplicativo host.

Substitua o c?digo no arquivo Home.js pelo c?digo a seguir. O c?digo mostra como voc? pode alterar as cadeias de caracteres usadas nos elementos da interface do usu?rio no Home.html com base no idioma de exibi??o do aplicativo host ou no idioma de edi??o do aplicativo host.

> **OBSERVA??O:** para alternar entre a altera??o da localiza??o do suplemento com base no idioma usado para edi??o, remova o coment?rio da linha de c?digo `var myLanguage = Office.context.contentLanguage;` e inclua o coment?rio na linha de c?digo `var myLanguage = Office.context.displayLanguage;`

```js
/// <reference path="../App.js" />
/// <reference path="../UIStrings.js" />


(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason)
    {
       
        $(document).ready(function () {
            app.initialize();

            // Get the language setting for editing document content.
            // To test this, uncomment the following line and then comment out the
            // line that uses Office.context.displayLanguage.
            // var myLanguage = Office.context.contentLanguage;

            // Get the language setting for UI display in the host application.
            var myLanguage = Office.context.displayLanguage;            
            var UIText;

            // Get the resource strings that match the language.
            // Use the UIStrings object from the UIStrings.js file
            // to get the JSON object with the correct localized strings.
            UIText = UIStrings.getLocaleStrings(myLanguage);            

            // Set localized text for UI elements.
            $("#greeting").text(UIText.Greeting);
            $("#about").text(UIText.Instruction);
        });
    };    
})();
```

### <a name="test-your-localized-add-in"></a>Testar seu suplemento localizado

Para testar seu suplemento localizado, altere o idioma usado para exibir ou editar no aplicativo host e execute o seu suplemento. 

Para alterar o idioma usado para exibir ou editar no seu suplemento:

1. No Word 2013, selecione **Arquivo** > , **Op??es** > , **Idioma**. A figura a seguir mostra a caixa de di?logo **Op??es do Word** aberta na guia Idioma.
    
    *Figura 2. Op??es de idioma na caixa de di?logo Op??es do Word 2013*

    ![Caixa de di?logo Op??es do Word 2013.](../images/office15-app-how-to-localize-fig04.png)

2. Em **Escolher idiomas de exibi??o e da ajuda**, selecione o idioma desejado para exibi??o, por exemplo, espanhol, e selecione a seta para cima para mover o idioma espanhol para a primeira posi??o na lista. Ou, para alterar o idioma usado para edi??o, em **Escolher idiomas de edi??o**, escolha o idioma que voc? deseja usar para edi??o, por exemplo, espanhol, e selecione **Definir como Padr?o**.
    
3. Escolha **OK** para confirmar sua sele??o e feche o Word.
    
Execute o suplemento de exemplo. O suplemento do painel de tarefas ? carregado no Word 2013 e as cadeias de caracteres na interface do usu?rio do suplemento s?o alteradas para corresponder ao idioma usado pelo aplicativo host, conforme mostrado na figura a seguir.


*Figura 3. Interface do usu?rio do suplemento com o texto localizado*

![Aplicativo com texto localizado da interface do usu?rio.](../images/office15-app-how-to-localize-fig05.png)

## <a name="see-also"></a>Confira tamb?m

- [Diretrizes de design para suplementos do Office](../design/add-in-design.md)    
- [Identificadores de idioma e valores da ID de OptionState no Office 2013](http://technet.microsoft.com/en-us/library/cc179219%28Office.15%29.aspx)

[DefaultLocale]:        https://dev.office.com/reference/add-ins/manifest/defaultlocale
[Descri??o]:          https://dev.office.com/reference/add-ins/manifest/description
[DisplayName]:          https://dev.office.com/reference/add-ins/manifest/displayname
[IconUrl]:              https://dev.office.com/reference/add-ins/manifest/iconurl
[HighResolutionIconUrl]:https://dev.office.com/reference/add-ins/manifest/highresolutioniconurl
[Recursos]:            https://dev.office.com/reference/add-ins/manifest/resources
[SourceLocation]:       https://dev.office.com/reference/add-ins/manifest/sourcelocation
[Substitui??o]:             https://dev.office.com/reference/add-ins/manifest/override
[DesktopSettings]:      https://dev.office.com/reference/add-ins/manifest/desktopsettings
[TabletSettings]:       https://dev.office.com/reference/add-ins/manifest/tabletsettings
[PhoneSettings]:        https://dev.office.com/reference/add-ins/manifest/phonesettings
[displayLanguage]:  https://dev.office.com/reference/add-ins/shared/office.context.displaylanguage 
[contentLanguage]:  https://dev.office.com/reference/add-ins/shared/office.context.contentlanguage 
[RFC 3066]: https://www.rfc-editor.org/info/rfc3066
