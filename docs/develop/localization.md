---
title: Localização para Suplementos do Office
description: Você pode usar a API JavaScript para Office para determinar uma localidade e exibir sequências de caracteres com base na localidade do aplicativo host, ou interpretar, ou exibir dados com base na sua localidade.
ms.date: 01/23/2018
ms.openlocfilehash: 6271010a08266c71d0f8242acf22cc7b1c730381
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506053"
---
# <a name="localization-for-office-add-ins"></a>Localização para Suplementos do Office

Você pode implementar qualquer esquema de localização que seja apropriado para o seu suplemento do Office. A API JavaScript e o esquema do manifesto da plataforma de Suplementos do Office oferecem algumas opções. Você pode usar a API JavaScript para Office para determinar uma localidade e exibir sequências de caracteres com base na localidade do aplicativo host ou para interpretar ou exibir dados com base na localidade dos dados. Você pode usar o manifesto para especificar informações descritivas e local do arquivo do suplemento específicos por localidade. Como alternativa, você pode usar o script do Microsoft Ajax para dar suporte à globalização e localização.

## <a name="use-the-javascript-api-to-determine-locale-specific-strings"></a>Use a API JavaScript para determinar sequências de caracteres específicas por localidade

A API JavaScript para Office fornece duas propriedades que oferecem suporte à exibição ou interpretação de valores consistentes com a localidade do aplicativo host e dos dados:

- [Context.displayLanguage][displayLanguage] especifica a localidade (ou idioma) da interface do usuário do aplicativo host. O exemplo a seguir verifica se o aplicativo host usa a localidade en-US ou fr-FR e exibe uma saudação específica para a localidade.
    
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

- [Context.contentLanguage][contentLanguage] especifica a localidade (ou o idioma) dos dados. Estendendo o último exemplo de código, em vez de verificar a propriedade [displayLanguage], atribua `myLanguage` à propriedade [contentLanguage] e use o restante do mesmo código para exibir uma saudação com base na localidade dos dados:
    
    ```js
    var myLanguage = Office.context.contentLanguage;
    ```

## <a name="control-localization-from-the-manifest"></a>Controlar a localização a partir do manifesto


Cada Suplemento do Office especifica um elemento [DefaultLocale] e uma localidade em seu manifesto. Por padrão, a plataforma de Suplementos do Office e os aplicativos host do Office aplicam os valores dos elementos [Description], [DisplayName], [IconUrl], [HighResolutionIconUrl] e [SourceLocation] a todas as localidades. Como opção, você pode oferecer suporte a valores específicos para localidades específicas, especificando um elemento-filho [Override]para cada localidade adicional, para qualquer um desses cinco elementos. O valor do elemento [DefaultLocale] e do atributo `Locale` do elemento [Override] é especificado de acordo com a [RFC 3066], "Marcas para a Identificação de Idiomas". A Tabela 1 descreve o suporte de localização para esses elementos.

**Tabela 1. Suporte para localização**


|**Elemento**|**Suporte para localização**|
|:-----|:-----|
|[Description]   |Os usuários de cada localidade especificada podem ver uma descrição localizada do suplemento na AppSource (ou no catálogo privado).<br/>Para os suplementos do Outlook, os usuários podem ver a descrição no Centro de Administração do Exchange (EAC) após a instalação.|
|[DisplayName]   |Os usuários de cada localidade especificada podem ver uma descrição localizada do suplemento na AppSource (ou no catálogo privado).<br/>Para os suplementos do Outlook, os usuários podem ver o nome para exibição como um rótulo para o botão do suplemento do Outlook e no EAC após a instalação.<br/>Para suplementos de painel de tarefas e de conteúdo, os usuários podem ver o nome para exibição na faixa de opções após a instalação do suplemento.|
|[IconUrl]        |A imagem do ícone é opcional. Você pode usar a mesma técnica de substituição para especificar uma determinada imagem para uma cultura específica. Se você usar e localizar um ícone, os usuários em cada localidade que você especificar poderão ver uma imagem de ícone localizada para o suplemento.<br/>Para suplementos do Outlook, os usuários podem ver o ícone no EAC depois de instalar o suplemento.<br/>Para suplementos de painel de tarefas e de conteúdo, os usuários podem ver o ícone na faixa de opções após a instalação do suplemento.|
|[HighResolutionIconUrl] **Importante:** este elemento só está disponível se você usar a versão 1.1 do manifesto do suplemento.|A imagem do ícone em alta resolução é opcional, mas se for especificada, deverá ocorrer após o elemento [IconUrl]. Quando [HighResolutionIconUrl] for especificado e o suplemento estiver instalado em um dispositivo que ofereça suporte à alta resolução de dpi, o valor [HighResolutionIconUrl] será usado em vez do valor de [IconUrl].<br/>Você pode usar a mesma técnica de substituição para especificar uma determinada imagem para uma cultura específica. Se você usar e localizar um ícone, os usuários em cada localidade que você especificar poderão ver uma imagem de ícone localizada para o suplemento.<br/>Para suplementos do Outlook, os usuários podem ver o ícone no EAC depois de instalar o suplemento.<br/>Para suplementos de painel de tarefas e de conteúdo, os usuários podem ver o ícone na faixa de opções após a instalação do suplemento.|
|[Recursos] **Importante:** este elemento só está disponível se você usar a versão 1.1 do manifesto do suplemento.   |Os usuários em cada localidade especificada podem ver recursos de sequências de caracteres e ícones que você criou especificamente para o suplemento dessa localidade. |
|[SourceLocation]   |Os usuários em cada localidade especificada podem ver uma página da Web que você criou especificamente para o suplemento para essa localidade. |


> [!NOTE] 
> Você só pode localizar o nome para exibição e a descrição para localidades suportadas pelo Office. Confira [Identificadores de idioma e valores de OptionState Id no Office 2013](https://docs.microsoft.com/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15)) para obter uma lista de idiomas e localidades para a versão atual do Office.


### <a name="examples"></a>Exemplos

Por exemplo, um suplemento do Office pode especificar o [DefaultLocale] como `en-us`. Para o elemento [DisplayName], o suplemento pode especificar um elemento filho [Override] para a localidade `fr-fr`, como é mostrado abaixo. 


```xml
<DefaultLocale>en-us</DefaultLocale>
...
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

> [!NOTE] 
> Se for preciso localizar para mais de uma área dentro de uma família de idiomas, como `de-de` e `de-at`, recomendamos que você use elementos `Override` separados para cada área. Usar apenas o nome do idioma sozinho, `de` neste caso, não tem suporte em todas as combinações de plataformas e aplicativos de host do Office.

Isso significa que o suplemento pressupõe a localidade `en-us` como padrão. Os usuários veem o nome para exibição em inglês "Video player" para todas as localidades, a menos que a localidade do computador cliente seja `fr-fr`. Nesse caso os usuários verão o nome para exibição em francês "Lecteur vidéo".

> [!NOTE] 
> Você só pode especificar uma única substituição por idioma, inclusive para a localidade padrão. Por exemplo, se sua localidade padrão for `en-us`, você não pode especificar também uma substituição para `en-us`. 

O exemplo a seguir se aplica a uma substituição de localidade para o elemento [Description]. Primeiro especifica a localidade padrão `en-us` e uma descrição em inglês e, em seguida, especifica uma instrução [Override] com uma descrição em francês para a localidade `fr-fr`:

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

Isso significa que o suplemento pressupõe a localidade `en-us` como padrão. Os usuários verão a descrição em inglês no atributo `DefaultValue` para todas as localidades, a menos que a localidade do computador cliente seja `fr-fr`. Nesse caso, verão a descrição em francês.

No exemplo a seguir, o suplemento especifica uma imagem separada que é mais apropriada para a localidade e a cultura `fr-fr`. Os usuários verão a imagem DefaultLogo.png por padrão, exceto quando a localidade do computador cliente for `fr-fr`. Nesse caso, verão a imagem FrenchLogo.png. 


```xml
<!-- Replace "domain" with a real web server name and path. -->
<IconUrl DefaultValue="https://<domain>/DefaultLogo.png"/>
<Override Locale="fr-fr" Value="https://<domain>/FrenchLogo.png"/>
```

O exemplo a seguir mostra como localizar um recurso na seção `Resources`. Ele aplica uma substituição de localidade para uma imagem apropriada para a cultura `ja-jp`.

```xml
<Resources>
      <bt:Images>
        <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
        </bt:Image>
 ...
```


Para o elemento [SourceLocation], o suporte a localidades adicionais implica em fornecer um arquivo HTML distinto para cada uma das localidades especificadas. Os usuários de cada localidade que você especificar poderão ver uma página da Web personalizada que foi projetada para eles.

Para suplementos do Outlook, o elemento [SourceLocation] também é alinhado ao fator forma, o que permite que você forneça um arquivo HTML de origem localizado e distinto para cada fator forma correspondente. Você pode especificar um ou mais elementos [Override] filho em cada configuração aplicável ([DesktopSettings], [TabletSettings] ou [PhoneSettings]). O exemplo a seguir mostra os elementos de configuração para fatores forma de desktop, tablet e smartphone, cada um deles com um arquivo HTML para a localidade padrão e outro para a localidade em francês.


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

## <a name="match-datetime-format-with-client-locale"></a>Correspondência entre o formato de data/hora e a localidade do cliente

Você pode obter a localidade da interface do usuário do aplicativo host usando a propriedade [displayLanguage]. Assim pode exibir valores de data e hora em um formato consistente com a localidade atual do aplicativo host. Uma maneira de fazer isso é preparar um arquivo de recurso que especifica o formato de exibição de data/hora a ser usado para cada localidade suportada pelo seu suplemento do Office. Em tempo de execução, seu suplemento pode usar o arquivo de recurso e fazer a correspondência entre o formato de data/hora apropriado e a localidade obtida na propriedade [displayLanguage].

Você pode obter a localidade dos dados do aplicativo host usando a propriedade [contentLanguage]. Com base nesse valor, você pode interpretar ou exibir adequadamente as sequências de caracteres de data/hora. Por exemplo, a localidade `jp-JP` expressa valores de data/hora como `yyyy/MM/dd`, e a localidade `fr-FR` como `dd/MM/yyyy`.


## <a name="use-ajax-for-globalization-and-localization"></a>Use Ajax para globalização e localização


Se você usar o Visual Studio para criar suplementos do Office, o .NET Framework e o Ajax oferecem maneiras para globalizar e localizar arquivos de script de cliente.

Você pode globalizar e utilizar as extensões de tipo JavaScript de [Date](http://msdn.microsoft.com/library/caf98d32-2de2-4704-8198-692350343681.aspx) e [Number](http://msdn.microsoft.com/library/c216d3a1-12ae-47d1-bca1-c3666d04572f.aspx) e o objeto [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) no código JavaScript de um suplemento do Office para exibir valores com base nas configurações de localização do navegador atual. Para saber mais, confira [Passo a passo: como globalizar uma data usando o script de cliente](http://msdn.microsoft.com/library/69b34e6d-d590-4d03-a763-b7ae54b47d74.aspx).

Você pode incluir sequências de caracteres de recurso localizadas diretamente em arquivos JavaScript autônomos para fornecer arquivos de script de cliente para diferentes localidades, que estão definidos no navegador ou fornecidos pelo usuário. Crie um arquivo de script separado para cada localidade suportada. Em cada arquivo de script, inclua um objeto no formato JSON que contenha as sequências de caracteres de recursos para essa localidade. Os valores localizados serão aplicados quando o script for executado no navegador. 


## <a name="example-build-a-localized-office-add-in"></a>Exemplo: Criação de um suplemento do Office localizado

Esta seção fornece exemplos que mostram como localizar a descrição, o nome para exibição e a interface do usuário de um suplemento do Office.

Para executar o exemplo de código, configure o Microsoft Office 2013 em seu computador para usar idiomas adicionais de modo que você possa testar seu suplemento, alternando o idioma usado para exibição em menus e comandos para edição,  revisão de texto ou ambos.

Além disso, você precisará criar um projeto de suplemento do Office no Visual Studio 2015.

> [!NOTE] 
> Para baixar o Visual Studio 2015, confira a [página Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs). Esta página também tem um link para Office Developer Tools.

### <a name="configure-office-2013-to-use-additional-languages-for-display-or-editing"></a>Configure o Office 2013 para usar idiomas adicionais para exibição ou edição

Você pode usar um pacote de idiomas do Office 2013 para instalar um idioma adicional. Para saber mais sobre os pacotes de idioma e onde obtê-los, confira [Opções de idioma do Office 2013](http://office.microsoft.com/language-packs/).

> [!NOTE] 
> Se você for assinante do MSDN, é possível que já tenha os pacotes de idiomas do Office 2013. Para determinar se a sua assinatura permite baixar os pacotes de idiomas do Office 2013, acesse [Página Inicial de Assinaturas do MSDN](https://msdn.microsoft.com/subscriptions/manage/), insira Pacote de idiomas do Office 2013 em **Download de softwares**, escolha **Pesquisa**  e selecione **Produtos disponíveis com minha assinatura**. Em **Idioma**, marque a caixa de seleção do Pacote de idiomas que você deseja baixar e, em seguida, selecione **Ir**. 

Depois de instalar o pacote de idiomas, você pode configurar o Office 2013 para usar o idioma instalado para exibição da interface do usuário, para edição de conteúdo do documento, ou ambos. O exemplo neste artigo usa uma instalação do Office 2013 com o pacote de idiomas de espanhol aplicado.

### <a name="create-an-office-add-in-project"></a>Crie um projeto de suplemento do Office

1. No Visual Studio, escolha **Arquivo** > **Novo Projeto**.
    
2. Na caixa de diálogo **Novo Projeto**, em **Modelos**, expanda **Visual Basic** ou **Visual C#**, expanda **Office/SharePoint** e, em seguida, selecione **Suplementos do Office**.
    
3. Escolha **Suplemento do Office** e, em seguida, dê um nome para o seu suplemento, por exemplo, WorldReadyAddIn. Escolha **OK**.
    
4. Na caixa de diálogo **Criar Suplemento do Office**, selecione **Painel de tarefas** e selecione **Avançar**. Na próxima página, desmarque todos os aplicativos, exceto o Word. Selecione **Concluir** para criar o projeto.
    

### <a name="localize-the-text-used-in-your-add-in"></a>Localize o texto usado no seu suplemento

O texto que você deseja localizar em outro idioma aparece em duas áreas:

-  **Nome para exibição e descrição do suplemento**. São controlados por entradas no arquivo de manifesto do suplemento.
    
-  **Interface do usuário do suplemento**. Você pode localizar as sequências de caracteres que aparecem na interface do usuário do seu suplemento com código JavaScript, usando um arquivo de recurso separado que contenha as sequências de caracteres localizadas, por exemplo.
    
Para localizar o nome para exibição e a descrição do suplemento:

1. No **Gerenciador de Soluções**, expanda **WorldReadyAddIn**, **WorldReadyAddInManifest** e, em seguida, selecione **WorldReadyAddIn.xml**.
    
2. No WorldReadyAddInManifest.xml, substitua os elementos [DisplayName] e [Description] com o seguinte bloco de código:
    
    > [!NOTE] 
    > Você pode substituir as sequências de caracteres localizadas em espanhol usadas neste exemplo pelos elementos [DisplayName] e [Description] com as sequências de caracteres localizadas em qualquer outro idioma.

    ```xml
    <DisplayName DefaultValue="World Ready add-in">
      <Override Locale="es-es" Value="Aplicación de uso internacional"/>
    </DisplayName>
    <Description DefaultValue="An add-in for testing localization">
      <Override Locale="es-es" Value="Una aplicación para la prueba de la localización"/>
    </Description>
    ```

3. Quando você altera o idioma de exibição do Office 2013 do inglês para o espanhol, por exemplo, e executa o suplemento, o nome para exibição do suplemento e a descrição são mostrados com texto localizado. 
    
Para definir a interface do usuário do suplemento:

1. No Visual Studio, no **Gerenciador de Soluções**, selecione **Home.html**.
    
2. Substitua o código HTML em Home.html pelo seguinte.
    
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
    
A figura a seguir mostra o elemento de cabeçalho (h1) e o elemento de parágrafo (p) que exibirão o texto localizado quando seu suplemento de exemplo for executado.

*Figura 1. A interface do usuário do suplemento*

![Interface de usuário do aplicativo com seções destacadas](../images/office15-app-how-to-localize-fig03.png)

### <a name="add-the-resource-file-that-contains-the-localized-strings"></a>Adicione o arquivo de recurso que contém as sequências de caracteres localizadas

O arquivo de recurso JavaScript contém as sequências de caracteres usadas na interface do usuário do suplemento. A interface do usuário do suplemento de exemplo tem um elemento h1 que exibe uma saudação e um elemento p que apresenta o suplemento ao usuário. 

Para habilitar sequências de caracteres localizadas para o cabeçalho e o parágrafo, coloque-as em um arquivo de recurso separado. O arquivo de recurso cria um objeto JavaScript que contém um objeto JSON (JavaScript Object Notation) separado para cada conjunto de sequência de caracteres localizadas. O arquivo de recurso também fornece um método para obter o objeto JSON apropriado para uma determinada localidade. 

Para adicionar o arquivo de recurso ao projeto do suplemento:

1. No **Gerenciador de Soluções** no Visual Studio, escolha a pasta **Suplemento** no projeto da Web do suplemento de exemplo e selecione **Adicionar** > **Arquivo JavaScript**.
    
2. Na caixa de diálogo **Especificar o nome do item**, insira UIStrings.js.
    
3. Adicione o código a seguir ao arquivo UIStrings.js.

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

O arquivo de recurso UIStrings.js cria o objeto, **UIStrings**, que contém as sequências de caracteres localizadas para a interface do usuário do suplemento. 

### <a name="localize-the-text-used-for-the-add-in-ui"></a>Localize o texto usado na interface do usuário do suplemento

Para usar o arquivo de recurso no seu suplemento, você precisará adicionar uma marcação de script na Home.html. Quando a Home.html for carregada, o UIStrings.js será executado e o objeto **UIStrings** que você utiliza para obter a sequência de caracteres estará disponível para o seu código. Adicione o seguinte HTML à marcação head da Home.html para disponibilizar **UIStrings** para o seu código.

```html
<!-- Resource file for localized strings:                                                          -->
<script src="../UIStrings.js" type="text/javascript"></script>
```

Agora você pode usar o objeto **UIStrings** para definir as sequências de caracteres da interface do usuário do seu suplemento.

Se você quiser alterar a localização do seu suplemento com base no idioma usado para exibição de menus e comandos no aplicativo host, use a propriedade **Office.context.displayLanguage** para obter a localidade desse idioma. Por exemplo, se o idioma do aplicativo host utilizar espanhol para exibir menus e comandos, a propriedade **Office.context.displayLanguage** retornará o código es-ES.

Se você quiser alterar a localização do seu suplemento com base no idioma que está sendo usado para editar o conteúdo do documento, use a propriedade **Office.context.contentLanguage** para obter a localidade do idioma. Por exemplo, se o idioma do aplicativo host utilizar espanhol para editar o conteúdo do documento, a propriedade **Office.context.contentLanguage** retornará o código es-ES.

Depois que você souber o idioma que o aplicativo host está utilizando, é possível usar **UIStrings** para obter o conjunto de sequências de caracteres localizadas correspondente ao idioma do aplicativo host.

Substitua o código no arquivo Home.js pelo código a seguir. O código mostra como você pode alterar as sequências de caracteres usadas nos elementos da interface do usuário no Home.html com base no idioma de exibição do aplicativo host ou no idioma de edição do aplicativo host.

> [!NOTE] 
> Para alternar entre a alteração da localização do suplemento com base no idioma usado para edição, remova a marca de comentário da linha de código `var myLanguage = Office.context.contentLanguage;` e inclua o comentário fora da linha de código `var myLanguage = Office.context.displayLanguage;`

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

Para testar seu suplemento localizado, altere o idioma usado para exibição ou editação no aplicativo host e execute o seu suplemento. 

Para alterar o idioma usado para exibição ou edição no seu suplemento:

1. No Word 2013, selecione **Arquivo** > **Opções** > **Idioma**. A figura a seguir mostra a caixa de diálogo **Opções do Word** aberta na guia Idioma.
    
    *Figura 2. Opções de idioma na caixa de diálogo Opções do Word 2013*

    ![Caixa de diálogo Opções do Word 2013.](../images/office15-app-how-to-localize-fig04.png)

2. Em **Escolher idiomas para exibição e ajuda**, selecione o idioma desejado para exibição, por exemplo, espanhol, e selecione a seta para cima para mover o idioma espanhol para a primeira posição na lista. Ou, para alterar o idioma usado para edição, em **Escolher idiomas para edição**, escolha o idioma que você deseja usar para edição, por exemplo, espanhol, e selecione **Definir como padrão**.
    
3. Escolha **OK** para confirmar sua seleção e feche o Word.
    
Execute o suplemento de exemplo. O suplemento do painel de tarefas é carregado no Word 2013 e as sequências de caracteres na interface do usuário do suplemento são alteradas para corresponder ao idioma usado pelo aplicativo host, conforme mostrado na figura a seguir.


*Figura 3. Interface do usuário do suplemento com o texto localizado*

![Aplicativo com texto localizado na interface do usuário.](../images/office15-app-how-to-localize-fig05.png)

## <a name="see-also"></a>Confira também

- [Diretrizes de design para suplementos do Office](../design/add-in-design.md)    
- [Identificadores de idioma e valores de OptionState Id no Office 2013](https://docs.microsoft.com/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15))

[DefaultLocale]:        https://docs.microsoft.com/office/dev/add-ins/reference/manifest/defaultlocale?view=office-js
[Description]:          https://docs.microsoft.com/office/dev/add-ins/reference/manifest/description?view=office-js
[DisplayName]:          https://docs.microsoft.com/office/dev/add-ins/reference/manifest/displayname?view=office-js
[IconUrl]:              https://docs.microsoft.com/office/dev/add-ins/reference/manifest/iconurl?view=office-js
[HighResolutionIconUrl]:https://docs.microsoft.com/office/dev/add-ins/reference/manifest/highresolutioniconurl?view=office-js
[Recursos]:            https://docs.microsoft.com/office/dev/add-ins/reference/manifest/resources?view=office-js
[SourceLocation]:       https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation?view=office-js
[Substituição]:             https://docs.microsoft.com/office/dev/add-ins/reference/manifest/override?view=office-js
[DesktopSettings]:      https://docs.microsoft.com/office/dev/add-ins/reference/manifest/desktopsettings?view=office-js
[TabletSettings]:       https://docs.microsoft.com/office/dev/add-ins/reference/manifest/tabletsettings?view=office-js
[PhoneSettings]:        https://docs.microsoft.com/office/dev/add-ins/reference/manifest/phonesettings?view=office-js
[displayLanguage]:  https://docs.microsoft.com/javascript/api/office/office.context?view=office-js#displaylanguage 
[contentLanguage]:  https://docs.microsoft.com/javascript/api/office/office.context?view=office-js#contentlanguage 
[RFC 3066]: https://www.rfc-editor.org/info/rfc3066
