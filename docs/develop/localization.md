---
title: Localização para Suplementos do Office
description: Use Office API JavaScript do Office para determinar uma localidade e exibir cadeias de caracteres com base na localidade do aplicativo Office, ou para interpretar ou exibir dados com base na localidade dos dados.
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: fc9d6a9c36ceabf96ee484eb4879074be8eecef76a561570cbde08458671ee3c
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57080502"
---
# <a name="localization-for-office-add-ins"></a>Localização para Suplementos do Office

Você pode implementar qualquer esquema de localização que seja apropriado para o seu Suplemento do Office. A API JavaScript e o esquema do manifesto da plataforma de Suplementos do Office oferecem algumas opções. Você pode usar Office API JavaScript do Office para determinar uma localidade e exibir cadeias de caracteres com base na localidade do aplicativo Office, ou para interpretar ou exibir dados com base na localidade dos dados. Você pode usar o manifesto para especificar informações descritivas e o local do arquivo do suplemento específico da localidade. Como alternativa, você pode usar o script do Microsoft Ajax para dar suporte à globalização e localização.

## <a name="use-the-javascript-api-to-determine-locale-specific-strings"></a>Usar a API JavaScript para determinar cadeias de caracteres específicas da localidade

A OFFICE JavaScript fornece duas propriedades que suportam a exibição ou interpretação de valores consistentes com a localidade do aplicativo Office dados.

- [Context.displayLanguage][displayLanguage] especifica a localidade (ou idioma) da interface do usuário do aplicativo Office. O exemplo a seguir verifica se o aplicativo Office usa a localidade en-US ou fr-FR e exibe uma saudação específica da localidade.

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

- [Context.contentLanguage][contentLanguage] especifica a localidade (ou o idioma) dos dados. Estendendo o último exemplo de código, em vez de verificar a propriedade [displayLanguage,] atribua o valor da propriedade `myLanguage` [contentLanguage] e use o restante do mesmo código para exibir uma saudação com base na localidade dos dados.

    ```js
    var myLanguage = Office.context.contentLanguage;
    ```

## <a name="control-localization-from-the-manifest"></a>Controlar a localização do manifesto

Cada Suplemento do Office especifica um elemento [DefaultLocale] e uma localidade em seu manifesto. Por padrão, Office plataforma de Office aplicativos cliente aplicam os valores dos elementos [Description,] [DisplayName,] [IconUrl,] [HighResolutionIconUrl]e [SourceLocation] a todas as localidades. Como opção, você pode dar suporte a valores específicos para localidades específicas, especificando um elemento-filho [Override]para cada localidade adicional, para qualquer um desses cinco elementos. O valor do elemento [DefaultLocale] e do atributo `Locale` do elemento [Override] é especificado de acordo com o [RFC 3066], "Marcas para a Identificação dos Idiomas". A Tabela 1 descreve o suporte de localização para esses elementos.

*Tabela 1. Suporte de localização*

|**Elemento**|**Suporte de localização**|
|:-----|:-----|
|[Descrição]   |Os usuários de cada localidade especificada podem ver uma descrição localizada do suplemento no AppSource (ou no catálogo privado).<br/>Para os suplementos do Outlook, os usuários podem ver a descrição no Centro de Administração do Exchange (EAC) após a instalação.|
|[DisplayName]   |Os usuários de cada localidade especificada podem ver uma descrição localizada do suplemento no AppSource (ou no catálogo privado).<br/>Para os suplementos do Outlook, os usuários podem ver o nome de exibição como um rótulo para o botão de suplemento do Outlook e no EAC após a instalação.<br/>Para os suplementos do painel de tarefas e do conteúdo, os usuários podem ver o nome de exibição na faixa de opções após a instalação do suplemento.|
|[IconUrl]        |A imagem do ícone é opcional. Você pode usar a mesma técnica de substituição para especificar uma determinada imagem para uma cultura específica. Se você usar e localizar um ícone, os usuários em cada localidade que você especificar poderão ver uma imagem de ícone localizada para o suplemento.<br/>Para suplementos do Outlook, os usuários podem ver o ícone no EAC depois de instalar o suplemento.<br/>Para os suplementos do painel de tarefas e do conteúdo, os usuários podem ver o ícone na faixa de opções após a instalação do suplemento.|
|[HighResolutionIconUrl] **Importante:** este elemento só fica disponível ao usar a versão 1.1 do manifesto do suplemento.|A imagem do ícone de alta resolução é opcional, mas se ela for especificada, deverá ocorrer após o elemento [IconUrl]. Quando [HighResolutionIconUrl] for especificado e o suplemento estiver instalado em um dispositivo que ofereça suporte à resolução dpi alto, o valor [HighResolutionIconUrl] é usado em vez do valor para [IconUrl].<br/>Você pode usar a mesma técnica de substituição para especificar uma determinada imagem para uma cultura específica. Se você usar e localizar um ícone, os usuários em cada localidade que você especificar podem ver uma imagem de ícone localizada para o suplemento.<br/>Para suplementos do Outlook, os usuários podem ver o ícone no EAC depois de instalar o suplemento.<br/>Para os suplementos do painel de tarefas e do conteúdo, os usuários podem ver o ícone na faixa de opções após a instalação do suplemento.|
|[Recursos] **Importante:** este elemento só fica disponível ao usar a versão 1.1 do manifesto do suplemento.   |Os usuários em cada localidade especificada podem ver recursos de cadeias de caracteres e de ícones que você projetou especificamente para o suplemento dessa localidade. |
|[SourceLocation]   |Os usuários de cada localidade especificada podem ver a página da Web que você projetou especificamente para o suplemento dessa localidade. |

> [!NOTE]
> Você só pode localizar o nome de exibição e a descrição das localidades para as quais o Office oferece suporte. Consulte [Identificadores de idioma e valores de OptionState Id no Office 2013](/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15)) para obter uma lista de idiomas e localidades da versão atual do Office.

### <a name="examples"></a>Exemplos

Por exemplo, um Suplemento do Office pode especificar [DefaultLocale] como `en-us`. Para o elemento [DisplayName], o suplemento pode especificar um elemento filho [Override] para a localidade `fr-fr`, como mostrado abaixo.

```xml
<DefaultLocale>en-us</DefaultLocale>
...
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

> [!NOTE]
> Se for preciso localizar para mais de uma área dentro de uma família de idiomas, como `de-de` e `de-at`, recomendamos que você use elementos `Override` separados para cada área. Usar apenas o nome do idioma, nesse caso, não é suportado em todas as combinações de aplicativos e plataformas Office `de` cliente.

Isso significa que o suplemento pressupõe a localidade `en-us` como padrão. Os usuários veem o nome de exibição em inglês "Video player" para todas as localidades, a menos que a localidade do computador cliente fosse `fr-fr`, nesse caso os usuários veria o nome de exibição em francês "Lecteur vidéo".

> [!NOTE]
> Você só pode especificar uma única substituição por idioma, inclusive para a localidade padrão. Por exemplo, se sua localidade padrão for `en-us`, não é possível especificar também uma substituição para `en-us`.

O exemplo a seguir aplica uma substituição de localidade para o [elemento Description.] Primeiro especifica uma localidade padrão e uma descrição em inglês e, em seguida, especifica uma instrução Override com uma descrição em francês `en-us` para a [] `fr-fr` localidade.

```xml
<DefaultLocale>en-us</DefaultLocale>
...
<Description DefaultValue=
   "Watch YouTube videos referenced in the emails you receive
   without leaving your email client.">
   <Override Locale="fr-fr" Value=
   "Visualisez les vidéos YouTube référencées dans vos courriers 
   électronique directement depuis Outlook."/>
</Description>
```

Isso significa que o suplemento pressupõe a localidade `en-us` como padrão. Os usuários veriam a descrição em inglês no atributo `DefaultValue` para todas as localidades, a menos que a localidade do computador cliente fosse `fr-fr`, nesse caso, eles veriam a descrição em francês.

No exemplo a seguir, o suplemento especifica uma imagem separada mais apropriada para a localidade e a cultura `fr-fr`. Os usuários verão a imagem DefaultLogo.png por padrão, exceto quando a localidade do computador cliente for `fr-fr`. Nesse caso, os usuários veriam a imagem FrenchLogo.png.

```xml
<!-- Replace "domain" with a real web server name and path. -->
<IconUrl DefaultValue="https://<domain>/DefaultLogo.png"/>
<Override Locale="fr-fr" Value="https://<domain>/FrenchLogo.png"/>
```

O exemplo a seguir mostra como localizar um recurso na seção `Resources`. Ele aplica um substituto local para uma imagem que é mais apropriada para a cultura `ja-jp`.

```xml
<Resources>
      <bt:Images>
        <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
        </bt:Image>
 ...
```

Para o elemento [SourceLocation], o suporte a localidades adicionais significa fornecer um arquivo HTML de origem separado para cada um dos locais especificados. Os usuários de cada localidade que você especificar poderão ver uma página da Web personalizada que foi projetada para eles.

Para suplementos do Outlook, o elemento [SourceLocation] também atribui o fator forma, o que permite que você forneça um arquivo HTML de origem localizado e distinto para cada fator de foram correspondente. Você pode especificar um ou mais elementos filho [Override] em cada configuração aplicável ([DesktopSettings], [TabletSettings] ou [PhoneSettings]). O exemplo a seguir mostra os elementos de configurações para fatores de forma de desktop, tablet e smartphone, cada um com um arquivo HTML para a localidade padrão e outro para a localidade francesa.

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

## <a name="localize-extended-overrides"></a>Localize substituições estendidas

Alguns recursos de extensibilidade de Office, como atalhos de teclado, são configurados com arquivos JSON hospedados em seu servidor, em vez de com o manifesto XML do complemento. Esta seção pressupõe que você esteja familiarizado com substituições estendidas. Consulte [Trabalhar com substituições estendidas do manifesto](extended-overrides.md) e do elemento [ExtendedOverrides.](../reference/manifest/extendedoverrides.md)

Use o `ResourceUrl` atributo do [elemento ExtendedOverrides](../reference/manifest/extendedoverrides.md) para apontar Office para um arquivo de recursos localizados. Apresentamos um exemplo a seguir.

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json" 
                       ResourceUrl="https://contoso.com/addin/my-resources.json">
    </ExtendedOverrides>
</OfficeApp>
```

O arquivo de substituições estendidas usa tokens em vez de cadeias de caracteres. As cadeias de caracteres de nomes de tokens no arquivo de recursos. Veja a seguir um exemplo que atribui um atalho de teclado a uma função (definida em outro lugar) que exibe o painel de tarefas do complemento. Observação sobre essa marcação:

- O exemplo não é muito válido. (Adicionamos uma propriedade adicional necessária a ela abaixo.)
- Os tokens devem ter o formato **${resource.*name-of-resource*}**.

```json
{
    "actions": [
        {
            "id": "SHOWTASKPANE",
            "type": "ExecuteFunction",
            "name": "${resource.SHOWTASKPANE_action_name}"
        }
    ],
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "${resource.SHOWTASKPANE_default_shortcut}"
            }
        }
    ] 
}
```

O arquivo de recursos, que também é formatado por JSON, tem uma propriedade de nível superior que é dividida em `resources` subpropropriedades por localidade. Para cada localidade, uma cadeia de caracteres é atribuída a cada token que foi usado no arquivo de substituições estendidas. A seguir está um exemplo que tem cadeias de caracteres `en-us` para `fr-fr` e . Neste exemplo, o atalho de teclado é o mesmo em ambas as localidades, mas isso nem sempre será o caso, especialmente quando você estiver localizando para localidades que têm um alfabeto ou um sistema de escrita diferente e, portanto, um teclado diferente.

```json
{
    "resources":{ 
        "en-us": { 
            "SHOWTASKPANE_default_shortcut": { 
                "value": "CTRL+SHIFT+A", 
            }, 
            "SHOWTASKPANE_action_name": {
                "value": "Show task pane for add-in",
            }, 
        },
        "fr-fr": { 
            "SHOWTASKPANE_default_shortcut": { 
                "value": "CTRL+SHIFT+A", 
            }, 
            "SHOWTASKPANE_action_name": {
                "value": "Afficher le volet de tâche pour add-in",
              } 
        }
    }
}
```

Não há `default` nenhuma propriedade no arquivo que seja um ponto para as `en-us` `fr-fr` seções e. Isso acontece porque as cadeias de caracteres padrão, que são usadas quando a localidade do aplicativo host do Office não corresponder a nenhuma das propriedades *ll-cc* no arquivo de recursos, devem ser *definidas* no próprio arquivo de substituições estendido. Definir as cadeias de caracteres padrão diretamente no arquivo de substituições estendidas garante que o Office não baixe o arquivo de recurso quando a localidade do aplicativo Office corresponde à localidade padrão do complemento (conforme especificado no manifesto). A seguir está uma versão corrigida do exemplo anterior de um arquivo de substituições estendido que usa tokens de recurso.

```json
{
    "actions": [
        {
            "id": "SHOWTASKPANE",
            "type": "ExecuteFunction",
            "name": "${resource.SHOWTASKPANE_action_name}"
        }
    ],
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "${resource.SHOWTASKPANE_default_shortcut}"
            }
        }
    ],
    "resources": { 
        "default": { 
            "SHOWTASKPANE_default_shortcut": { 
                "value": "CTRL+SHIFT+A", 
            }, 
            "SHOWTASKPANE_action_name": {
                "value": "Show task pane for add-in",
            } 
        }
    }
}
```

## <a name="match-datetime-format-with-client-locale"></a>Fazer a correspondência entre o formato de data/hora e a localidade do cliente

Você pode obter a localidade da interface do usuário do aplicativo Office cliente usando a **[propriedade displayLanguage.]** Em seguida, você pode exibir valores de data e hora em um formato consistente com a localidade atual do Office aplicativo. Uma maneira de fazer isso é preparar um arquivo de recurso que especifica o formato de exibição de data/hora a ser usado em cada localidade com suporte do seu Suplemento do Office. Em tempo de executar, o seu complemento pode usar o arquivo de recurso e corresponder ao formato de data/hora apropriado com a localidade obtida da **[propriedade displayLanguage.]**

Você pode obter a localidade dos dados do aplicativo Office cliente usando a [propriedade contentLanguage.] Com base nesse valor, você pode, então, interpretar ou exibir adequadamente as cadeias de caracteres de data/hora. Por exemplo, a localidade `jp-JP` expressa valores de data/hora como `yyyy/MM/dd`, e a localidade `fr-FR` como `dd/MM/yyyy`.

## <a name="use-ajax-for-globalization-and-localization"></a>Usar o Ajax para a globalização e a localização

Se você usar o Visual Studio para criar Suplementos do Office, o .NET Framework e Ajax fornecem maneiras de globalizar e localizar arquivos de script de cliente.

Você pode globalizar e utilizar as extensões do tipo JavaScript de [Data](/previous-versions/bb310850(v=vs.140)) e [Número](/previous-versions/bb310835(v=vs.140)) e o objeto [Data](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) do JavaScript no código do JavaScript para um suplemento do Office para exibir valores com base nas configurações de localização do navegador atual. Para saber mais, confira [Passo a passo: como globalizar uma data usando o script de cliente](/previous-versions/bb386581(v=vs.140)).

Você pode incluir cadeias de caracteres de recurso localizadas diretamente em arquivos de JavaScript autônomos para fornecer arquivos de script de cliente para diferentes locais, que são definidos no navegador ou fornecidos pelo usuário. Crie um arquivo de script separado para cada localidade com suporte. Em cada arquivo de script, inclua um objeto no formato JSON que contenha as cadeias de caracteres de recursos para essa localidade. Os valores localizados serão aplicados quando o script for executado no navegador.

## <a name="example-build-a-localized-office-add-in"></a>Exemplo: Criar um Suplemento do Office localizado

Esta seção fornece exemplos que mostram como localizar uma descrição do Suplemento do Office, o nome de exibição e interface do usuário.

> [!NOTE]
> Para baixar Visual Studio 2019, consulte a [página Visual Studio IDE .](https://visualstudio.microsoft.com/vs/) Durante a instalação, você precisará selecionar a carga de trabalho de desenvolvimento do Office/SharePoint.

### <a name="configure-office-to-use-additional-languages-for-display-or-editing"></a>Configurar o Office para usar idiomas adicionais para exibição ou edição

Para executar o código de exemplo fornecido, configure o Office em seu computador para usar idiomas adicionais para que você possa testar seu complemento alternando o idioma usado para exibição em menus e comandos, para edição e revisão de texto, ou ambos.

Você pode usar um Office Language Pack para instalar um idioma adicional. Para saber mais sobre os Pacotes de Idiomas e onde obtê-los, veja [Language Accessory Pack do Office](https://support.microsoft.com/office/82ee1236-0f9a-45ee-9c72-05b026ee809f).

Depois de instalar o Language Accessory Pack, você pode configurar o Office para usar o idioma instalado para exibir na interface do usuário, para edição de conteúdo do documento ou ambos. O exemplo neste artigo usa uma instalação do Office que tenha o Pacote de Idiomas de espanhol aplicado.

### <a name="create-an-office-add-in-project"></a>Criar um projeto de Suplemento do Office

Você precisará criar um projeto de Visual Studio de Office 2019.

> [!NOTE]
> Se você ainda não instalou o Visual Studio 2019, consulte a página Visual Studio [IDE](https://visualstudio.microsoft.com/vs/) para obter instruções de download. Durante a instalação, você precisará selecionar a carga de trabalho de desenvolvimento do Office/SharePoint. Se você tiver instalado o Visual Studio 2019, [](/visualstudio/install/modify-visual-studio/) use o Instalador do Visual Studio para garantir que Office/SharePoint de desenvolvimento seja instalada.

1. Escolha **Criar um novo projeto**.

2. Usando a caixa de pesquisa, insira **suplemento**. Escolha **Suplemento do Word Web**, em seguida, selecione **Próximo**.

3. Nomeia **seu projeto WorldReadyAddIn** e selecione **Criar**.

4. O Visual Studio cria uma solução, e os dois projetos dele são exibidos no **Gerenciador de Soluções**. O arquivo **Home.html** é aberto no Visual Studio.

### <a name="localize-the-text-used-in-your-add-in"></a>Localizar o texto usado no seu suplemento

O texto que você deseja localizado para outro idioma aparece em duas áreas.

- **Nome de exibição e descrição do suplemento**. São controlados por entradas no arquivo de manifesto do suplemento.

- **Interface do usuário do suplemento**. Você pode localizar as cadeias de caracteres que aparecem na interface do usuário do seu suplemento usando códigos do JavaScript, por exemplo, usando um arquivo de recurso separado que contenha as cadeias de caracteres localizadas.

#### <a name="localize-the-add-in-display-name-and-description"></a>Localize o nome de exibição e a descrição do complemento

1. No **Explorador de** Soluções, **expanda WorldReadyAddIn,** **WorldReadyAddInManifest** e escolha **WorldReadyAddIn.xml**.

2. Em WorldReadyAddInManifest.xml, substitua os [elementos DisplayName] e [Description] pelo seguinte bloco de código.

    > [!NOTE]
    > Você pode substituir as cadeias de caracteres do idioma espanhol localizado usadas neste exemplo dos elementos [DisplayName] e [Description] pelas cadeias de caracteres localizados para qualquer outro idioma.

    ```xml
    <DisplayName DefaultValue="World Ready add-in">
      <Override Locale="es-es" Value="Aplicación de uso internacional"/>
    </DisplayName>
    <Description DefaultValue="An add-in for testing localization">
      <Override Locale="es-es" Value="Una aplicación para la prueba de la localización"/>
    </Description>
    ```

3. Quando você altera o idioma de exibição do Office 2013 do inglês para o espanhol, por exemplo, e executa o suplemento, o nome de exibição do suplemento e a descrição são mostrados com texto localizado.

#### <a name="lay-out-the-add-in-ui"></a>Desmarcar a interface do usuário do complemento

1. No Visual Studio, no **Gerenciador de Soluções**, selecione **Home.html**.

2. Substitua o conteúdo do elemento `<body>` no Home.html com o HTML a seguir e salve o arquivo.

    ```html
    <body>
        <!-- Page content -->
        <div id="content-header" class="ms-bgColor-themePrimary ms-font-xl">
            <div class="padding">
                <h1 id="greeting" class="ms-fontColor-white"></h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <div class="ms-font-m">
                    <p id="about"></p>
                </div>
            </div>
        </div>
    </body>
    ```

A figura a seguir mostra o elemento do cabeçalho (h1) e o elemento do parágrafo (p) que exibirá o texto localizado quando concluir as etapas restantes e executar o suplemento.

*Figura 1. A interface do usuário do suplemento*

![Interface do usuário do aplicativo com seções destacadas.](../images/office15-app-how-to-localize-fig03.png)

### <a name="add-the-resource-file-that-contains-the-localized-strings"></a>Adicionar o arquivo de recurso que contém as cadeias de caracteres localizadas

O arquivo de recurso do JavaScript contém as cadeias de caracteres usadas para a interface do usuário do suplemento. O HTML da interface do usuário do suplemento de amostra tem um elemento `<h1>` que exibe uma saudação e um elemento `<p>` que apresenta o suplemento ao usuário.

Para habilitar cadeias de caracteres para o cabeçalho e parágrafo, coloque as cadeias de caracteres em um arquivo de recurso separado. O arquivo de recurso cria um objeto do JavaScript que contém um objeto JSON (JavaScript Object Notation) separado para cada conjunto de cadeias de caracteres localizadas. O arquivo de recurso também fornece um método para obter o objeto JSON apropriado de volta para uma determinada localidade.

### <a name="add-the-resource-file-to-the-add-in-project"></a>Adicionar o arquivo de recurso ao projeto do add-in

1. No **Gerenciador de Soluções** no Visual Studio, clique com o botão direito no projeto **WorldReadyAddInWeb** e escolha **Adicionar** > **Novo Item**.

2. Na caixa de diálogo **Adicionar Novo Item**, escolha **Arquivo JavaScript**.

3. Insira **UIStrings.js** como nome do arquivo e escolha **Adicionar**.

4. Adicione o código a seguir ao arquivo UIStrings.js e salve-o.

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

O arquivo de recurso UIStrings.js cria o objeto, **UIStrings**, que contém as cadeias de caracteres localizadas para a interface do usuário do suplemento.

### <a name="localize-the-text-used-for-the-add-in-ui"></a>Localizar o texto usado na interface do usuário do suplemento

Para usar o arquivo de recurso no seu suplemento, você precisará adicionar a ele uma marca de script em Home.html. Quando Home.html for carregado, o UIStrings.js será executado e o objeto **UIStrings** que você utiliza para obter a cadeia de caracteres ficará disponível para seu código. Adicione o seguinte HTML à marca de cabeçalho do Home.html para tornar **UIStrings** disponível para seu código.

```html
<!-- Resource file for localized strings: -->
<script src="../UIStrings.js" type="text/javascript"></script>
```

Agora você pode usar o objeto **UIStrings** para definir as cadeias de caracteres da interface do usuário do seu suplemento.

Se você quiser alterar a localização do seu complemento com base em qual idioma é usado para exibição em menus e comandos no aplicativo cliente do Office, use a **propriedade Office.context.displayLanguage** para obter a localidade desse idioma. Por exemplo, se o idioma do aplicativo usar espanhol para exibição em menus e comandos, a propriedade **Office.context.displayLanguage** retornará o código de idioma es-ES.

Se você quiser alterar a localização do seu complemento com base em qual idioma está sendo usado para editar o conteúdo do documento, use a propriedade **Office.context.contentLanguage** para obter a localidade desse idioma. Por exemplo, se o idioma do aplicativo usar espanhol para editar o conteúdo do documento, a propriedade **Office.context.contentLanguage** retornará o código de idioma es-ES.

Depois de saber o idioma que o aplicativo está usando, você pode usar **UIStrings** para obter o conjunto de cadeias de caracteres localizadas que corresponde ao idioma do aplicativo.

Substitua o código no arquivo Home.js pelo código a seguir. O código mostra como você pode alterar as cadeias de caracteres usadas nos elementos da interface do usuário no Home.html com base no idioma de exibição do aplicativo ou no idioma de edição do aplicativo.

> [!NOTE]
> Para alternar entre a alteração da localização do suplemento com base no idioma usado para edição, remova o comentário da linha de código `var myLanguage = Office.context.contentLanguage;` e inclua o comentário na linha de código `var myLanguage = Office.context.displayLanguage;`

```js
/// <reference path="../App.js" />
/// <reference path="../UIStrings.js" />


(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason)
    {

        $(document).ready(function () {
            // Get the language setting for editing document content.
            // To test this, uncomment the following line and then comment out the
            // line that uses Office.context.displayLanguage.
            // var myLanguage = Office.context.contentLanguage;

            // Get the language setting for UI display in the Office application.
            var myLanguage = Office.context.displayLanguage;
            var UIText;

            // Get the resource strings that match the language.
            // Use the UIStrings object from the UIStrings.js file
            // to get the JSON object with the correct localized strings.
            UIText = UIStrings.getLocaleStrings(myLanguage);

            // Set localized text for UI elements.
            $("#greeting").text(UIText.Greeting);
            $("#about").text(UIText.Introduction);
        });
    };
})();
```

### <a name="test-your-localized-add-in"></a>Testar seu suplemento localizado

Para testar seu complemento localizado, altere o idioma usado para exibição ou edição no aplicativo Office e execute o seu complemento.

1. No Word, escolha **Arquivo** > **Opções** > **Idioma**. A figura a seguir mostra a caixa de diálogo **Opções do Word** aberta na guia Idioma.

    *Figura 2. Opções de idioma na caixa de diálogo Opções do Word*

    ![Caixa de diálogo Opções do Word.](../images/office15-app-how-to-localize-fig04.png)

2. Em **Escolher Idioma de Exibição**, selecione o idioma desejado para exibição, por exemplo, espanhol, e selecione a seta para cima para mover o idioma espanhol para a primeira posição na lista. Como alternativa, para alterar o idioma usado para edição, em **Escolher** Idiomas de Edição, escolha o idioma que você deseja usar para edição, por exemplo, espanhol e escolha **Definir como Padrão**.

3. Escolha **OK** para confirmar sua seleção e feche o Word.

4. Pressione **F5** no Visual Studio para executar o suplemento de amostra ou escolha **Depurar** > **Iniciar Depuração** na barra de menus.

5. No Word, escolha **Página Inicial** > **Mostrar Painel de Tarefas**.

Depois de executar, as cadeias de caracteres na interface do usuário do complemento mudam para corresponder ao idioma usado pelo aplicativo, conforme mostrado na figura a seguir.

*Figura 3. Interface do usuário do suplemento com o texto localizado*

![Aplicativo com texto localizado da IU.](../images/office15-app-how-to-localize-fig05.png)

## <a name="see-also"></a>Confira também

- [Diretrizes de design para suplementos do Office](../design/add-in-design.md)
- [Identificadores de idioma e valores da ID de OptionState no Office 2013](/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15))

[DefaultLocale]:         ../reference/manifest/defaultlocale.md
[Descrição]:           ../reference/manifest/description.md
[DisplayName]:           ../reference/manifest/displayname.md
[IconUrl]:               ../reference/manifest/iconurl.md
[HighResolutionIconUrl]: ../reference/manifest/highresolutioniconurl.md
[Resources]:             ../reference/manifest/resources.md
[SourceLocation]:        ../reference/manifest/sourcelocation.md
[Override]:              ../reference/manifest/override.md
[DesktopSettings]:       ../reference/manifest/desktopsettings.md
[TabletSettings]:        ../reference/manifest/tabletsettings.md
[PhoneSettings]:         ../reference/manifest/phonesettings.md
[displayLanguage]:       /javascript/api/office/office.context#displayLanguage
[contentLanguage]:       /javascript/api/office/office.context#contentLanguage
[RFC 3066]:              https://www.rfc-editor.org/info/rfc3066
