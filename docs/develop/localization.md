---
title: Localização para Suplementos do Office
description: Use a API JavaScript do Office para determinar uma localidade e exibir cadeias de caracteres com base na localidade do aplicativo do Office ou para interpretar ou exibir dados com base na localidade dos dados.
ms.date: 02/23/2021
localization_priority: Normal
ms.openlocfilehash: 8125bd55ce1d9dfe8e80bc4d80230555ec649787
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505273"
---
# <a name="localization-for-office-add-ins"></a><span data-ttu-id="5cb9a-103">Localização para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="5cb9a-103">Localization for Office Add-ins</span></span>

<span data-ttu-id="5cb9a-104">Você pode implementar qualquer esquema de localização que seja apropriado para o seu Suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-104">You can implement any localization scheme that's appropriate for your Office Add-in.</span></span> <span data-ttu-id="5cb9a-105">A API JavaScript e o esquema do manifesto da plataforma de Suplementos do Office oferecem algumas opções.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-105">The JavaScript API and manifest schema of the Office Add-ins platform provide some choices.</span></span> <span data-ttu-id="5cb9a-106">Você pode usar a API JavaScript do Office para determinar uma localidade e exibir cadeias de caracteres com base na localidade do aplicativo do Office ou para interpretar ou exibir dados com base na localidade dos dados.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-106">You can use the Office JavaScript API to determine a locale and display strings based on the locale of the Office application, or to interpret or display data based on the locale of the data.</span></span> <span data-ttu-id="5cb9a-107">Você pode usar o manifesto para especificar informações descritivas e o local do arquivo do suplemento específico da localidade.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-107">You can use the manifest to specify locale-specific add-in file location and descriptive information.</span></span> <span data-ttu-id="5cb9a-108">Como alternativa, você pode usar o script do Microsoft Ajax para dar suporte à globalização e localização.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-108">Alternatively, you can use Microsoft Ajax script to support globalization and localization.</span></span>

## <a name="use-the-javascript-api-to-determine-locale-specific-strings"></a><span data-ttu-id="5cb9a-109">Usar a API JavaScript para determinar cadeias de caracteres específicas da localidade</span><span class="sxs-lookup"><span data-stu-id="5cb9a-109">Use the JavaScript API to determine locale-specific strings</span></span>

<span data-ttu-id="5cb9a-110">A API JavaScript do Office fornece duas propriedades que suportam a exibição ou interpretação de valores consistentes com a localidade do aplicativo e dados do Office:</span><span class="sxs-lookup"><span data-stu-id="5cb9a-110">The Office JavaScript API provides two properties that support displaying or interpreting values consistent with the locale of the Office application and data:</span></span>

- <span data-ttu-id="5cb9a-111">[Context.displayLanguage][displayLanguage] especifica a localidade (ou idioma) da interface do usuário do aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-111">[Context.displayLanguage][displayLanguage] specifies the locale (or language) of the user interface of the Office application.</span></span> <span data-ttu-id="5cb9a-112">O exemplo a seguir verifica se o aplicativo do Office usa a localidade en-US ou fr-FR e exibe uma saudação específica da localidade.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-112">The following example verifies if the Office application uses the en-US or fr-FR locale, and displays a locale-specific greeting.</span></span>

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

- <span data-ttu-id="5cb9a-p103">[Context.contentLanguage][contentLanguage] especifica a localidade (ou o idioma) dos dados. Estendendo o último exemplo de código, em vez de verificar a propriedade [displayLanguage], atribua a `myLanguage` o valor da propriedade [contentLanguage] e use o restante do mesmo código para exibir uma saudação com base na localidade dos dados:</span><span class="sxs-lookup"><span data-stu-id="5cb9a-p103">[Context.contentLanguage][contentLanguage] specifies the locale (or language) of the data. Extending the last code sample, instead of checking the [displayLanguage] property, assign `myLanguage` the value of the [contentLanguage] property, and use the rest of the same code to display a greeting based on the locale of the data:</span></span>

    ```js
    var myLanguage = Office.context.contentLanguage;
    ```

## <a name="control-localization-from-the-manifest"></a><span data-ttu-id="5cb9a-115">Controlar a localização do manifesto</span><span class="sxs-lookup"><span data-stu-id="5cb9a-115">Control localization from the manifest</span></span>


<span data-ttu-id="5cb9a-116">Cada Suplemento do Office especifica um elemento [DefaultLocale] e uma localidade em seu manifesto.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-116">Every Office Add-in specifies a [DefaultLocale] element and a locale in its manifest.</span></span> <span data-ttu-id="5cb9a-117">Por padrão, a plataforma de Complementos do Office e os aplicativos cliente do Office aplicam os valores dos elementos [Description], [DisplayName,] [IconUrl,] [HighResolutionIconUrl]e [SourceLocation] a todas as localidades.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-117">By default, the Office Add-in platform and Office client applications apply the values of the [Description], [DisplayName], [IconUrl], [HighResolutionIconUrl], and [SourceLocation] elements to all locales.</span></span> <span data-ttu-id="5cb9a-118">Como opção, você pode dar suporte a valores específicos para localidades específicas, especificando um elemento-filho [Override]para cada localidade adicional, para qualquer um desses cinco elementos.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-118">You can optionally support specific values for specific locales, by specifying an [Override] child element for each additional locale, for any of these five elements.</span></span> <span data-ttu-id="5cb9a-119">O valor do elemento [DefaultLocale] e do atributo `Locale` do elemento [Override] é especificado de acordo com o [RFC 3066], "Marcas para a Identificação dos Idiomas".</span><span class="sxs-lookup"><span data-stu-id="5cb9a-119">The value for the [DefaultLocale] element and for the `Locale` attribute of the [Override] element is specified according to [RFC 3066], "Tags for the Identification of Languages."</span></span> <span data-ttu-id="5cb9a-120">A Tabela 1 descreve o suporte de localização para esses elementos.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-120">Table 1 describes the localizing support for these elements.</span></span>

<span data-ttu-id="5cb9a-121">*Tabela 1. Suporte de localização*</span><span class="sxs-lookup"><span data-stu-id="5cb9a-121">*Table 1. Localization support*</span></span>


|<span data-ttu-id="5cb9a-122">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="5cb9a-122">**Element**</span></span>|<span data-ttu-id="5cb9a-123">**Suporte de localização**</span><span class="sxs-lookup"><span data-stu-id="5cb9a-123">**Localization support**</span></span>|
|:-----|:-----|
|<span data-ttu-id="5cb9a-124">[Descrição]</span><span class="sxs-lookup"><span data-stu-id="5cb9a-124">[Description]</span></span>   |<span data-ttu-id="5cb9a-125">Os usuários de cada localidade especificada podem ver uma descrição localizada do suplemento no AppSource (ou no catálogo privado).</span><span class="sxs-lookup"><span data-stu-id="5cb9a-125">Users in each locale you specify can see a localized description for the add-in in AppSource (or private catalog).</span></span><br/><span data-ttu-id="5cb9a-126">Para os suplementos do Outlook, os usuários podem ver a descrição no Centro de Administração do Exchange (EAC) após a instalação.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-126">For Outlook add-ins, users can see the description in the Exchange Admin Center (EAC) after installation.</span></span>|
|<span data-ttu-id="5cb9a-127">[DisplayName]</span><span class="sxs-lookup"><span data-stu-id="5cb9a-127">[DisplayName]</span></span>   |<span data-ttu-id="5cb9a-128">Os usuários de cada localidade especificada podem ver uma descrição localizada do suplemento no AppSource (ou no catálogo privado).</span><span class="sxs-lookup"><span data-stu-id="5cb9a-128">Users in each locale you specify can see a localized description for the add-in in AppSource (or private catalog).</span></span><br/><span data-ttu-id="5cb9a-129">Para os suplementos do Outlook, os usuários podem ver o nome de exibição como um rótulo para o botão de suplemento do Outlook e no EAC após a instalação.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-129">For Outlook add-ins, users can see the display name as a label for the Outlook add-in button and in the EAC after installation.</span></span><br/><span data-ttu-id="5cb9a-130">Para os suplementos do painel de tarefas e do conteúdo, os usuários podem ver o nome de exibição na faixa de opções após a instalação do suplemento.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-130">For content and task pane add-ins, users can see the display name in the ribbon after installing the add-in.</span></span>|
|<span data-ttu-id="5cb9a-131">[IconUrl]</span><span class="sxs-lookup"><span data-stu-id="5cb9a-131">[IconUrl]</span></span>        |<span data-ttu-id="5cb9a-p105">A imagem do ícone é opcional. Você pode usar a mesma técnica de substituição para especificar uma determinada imagem para uma cultura específica. Se você usar e localizar um ícone, os usuários em cada localidade que você especificar poderão ver uma imagem de ícone localizada para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-p105">The icon image is optional. You can use the same override technique to specify a certain image for a specific culture. If you use and localize an icon, users in each locale you specify can see a localized icon image for the add-in.</span></span><br/><span data-ttu-id="5cb9a-135">Para suplementos do Outlook, os usuários podem ver o ícone no EAC depois de instalar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-135">For Outlook add-ins, users can see the icon in the EAC after installing the add-in.</span></span><br/><span data-ttu-id="5cb9a-136">Para os suplementos do painel de tarefas e do conteúdo, os usuários podem ver o ícone na faixa de opções após a instalação do suplemento.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-136">For content and task pane add-ins, users can see the icon in the ribbon after installing the add-in.</span></span>|
|<span data-ttu-id="5cb9a-137">[HighResolutionIconUrl] **Importante:** este elemento só fica disponível ao usar a versão 1.1 do manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-137">[HighResolutionIconUrl] **Important:** This element is available only when using add-in manifest version 1.1.</span></span>|<span data-ttu-id="5cb9a-p106">A imagem do ícone de alta resolução é opcional, mas se ela for especificada, deverá ocorrer após o elemento [IconUrl]. Quando [HighResolutionIconUrl] for especificado e o suplemento estiver instalado em um dispositivo que ofereça suporte à resolução dpi alto, o valor [HighResolutionIconUrl] é usado em vez do valor para [IconUrl].</span><span class="sxs-lookup"><span data-stu-id="5cb9a-p106">The high resolution icon image is optional but if it is specified, it must occur after the  [IconUrl] element. When [HighResolutionIconUrl] is specified, and the add-in is installed on a device that supports high dpi resolution, the [HighResolutionIconUrl] value is used instead of the value for [IconUrl].</span></span><br/><span data-ttu-id="5cb9a-p107">Você pode usar a mesma técnica de substituição para especificar uma determinada imagem para uma cultura específica. Se você usar e localizar um ícone, os usuários em cada localidade que você especificar podem ver uma imagem de ícone localizada para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-p107">You can use the same override technique to specify a certain image for a specific culture. If you use and localize an icon, users in each locale you specify can see a localized icon image for the add-in.</span></span><br/><span data-ttu-id="5cb9a-142">Para suplementos do Outlook, os usuários podem ver o ícone no EAC depois de instalar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-142">For Outlook add-ins, users can see the icon in the EAC after installing the add-in.</span></span><br/><span data-ttu-id="5cb9a-143">Para os suplementos do painel de tarefas e do conteúdo, os usuários podem ver o ícone na faixa de opções após a instalação do suplemento.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-143">For content and task pane add-ins, users can see the icon in the ribbon after installing the add-in.</span></span>|
|<span data-ttu-id="5cb9a-144">[Recursos] **Importante:** este elemento só fica disponível ao usar a versão 1.1 do manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-144">[Resources] **Important:** This element is available only when using add-in manifest version 1.1.</span></span>   |<span data-ttu-id="5cb9a-145">Os usuários em cada localidade especificada podem ver recursos de cadeias de caracteres e de ícones que você projetou especificamente para o suplemento dessa localidade.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-145">Users in each locale you specify can see string and icon resources that you specifically create for the add-in for that locale.</span></span> |
|<span data-ttu-id="5cb9a-146">[SourceLocation]</span><span class="sxs-lookup"><span data-stu-id="5cb9a-146">[SourceLocation]</span></span>   |<span data-ttu-id="5cb9a-147">Os usuários de cada localidade especificada podem ver a página da Web que você projetou especificamente para o suplemento dessa localidade.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-147">Users in each locale you specify can see a webpage that you specifically design for the add-in for that locale.</span></span> |


> [!NOTE]
> <span data-ttu-id="5cb9a-148">Você só pode localizar o nome de exibição e a descrição das localidades para as quais o Office oferece suporte.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-148">You can localize the description and display name for only the locales that Office supports.</span></span> <span data-ttu-id="5cb9a-149">Consulte [Identificadores de idioma e valores de OptionState Id no Office 2013](/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15)) para obter uma lista de idiomas e localidades da versão atual do Office.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-149">See [Language identifiers and OptionState Id values in Office 2013](/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15)) for a list of languages and locales for the current release of Office.</span></span>


### <a name="examples"></a><span data-ttu-id="5cb9a-150">Exemplos</span><span class="sxs-lookup"><span data-stu-id="5cb9a-150">Examples</span></span>

<span data-ttu-id="5cb9a-151">Por exemplo, um Suplemento do Office pode especificar [DefaultLocale] como `en-us`.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-151">For example, an Office Add-in can specify the [DefaultLocale] as `en-us`.</span></span> <span data-ttu-id="5cb9a-152">Para o elemento [DisplayName], o suplemento pode especificar um elemento filho [Override] para a localidade `fr-fr`, como mostrado abaixo.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-152">For the [DisplayName] element, the add-in can specify an [Override] child element for the locale `fr-fr`, as shown below.</span></span>


```xml
<DefaultLocale>en-us</DefaultLocale>
...
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

> [!NOTE]
> <span data-ttu-id="5cb9a-153">Se for preciso localizar para mais de uma área dentro de uma família de idiomas, como `de-de` e `de-at`, recomendamos que você use elementos `Override` separados para cada área.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-153">If you need to localize for more than one area within a language family, such as `de-de` and `de-at`, we recommend that you use separate `Override` elements for each area.</span></span> <span data-ttu-id="5cb9a-154">O uso apenas do nome de idioma, nesse caso, não é suportado em todas as combinações de `de` aplicativos cliente e plataformas do Office.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-154">Using just the language name alone, in this case, `de`, is not supported across all combinations of Office client applications and platforms.</span></span>

<span data-ttu-id="5cb9a-p111">Isso significa que o suplemento pressupõe a localidade `en-us` como padrão. Os usuários veem o nome de exibição em inglês "Video player" para todas as localidades, a menos que a localidade do computador cliente fosse `fr-fr`, nesse caso os usuários veria o nome de exibição em francês "Lecteur vidéo".</span><span class="sxs-lookup"><span data-stu-id="5cb9a-p111">This means that the add-in assumes the  `en-us` locale by default. Users see the English display name of "Video player" for all locales unless the client computer's locale is `fr-fr`, in which case users would see the French display name "Lecteur vidéo".</span></span>

> [!NOTE]
> <span data-ttu-id="5cb9a-157">Você só pode especificar uma única substituição por idioma, inclusive para a localidade padrão.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-157">You may only specify a single override per language, including for the default locale.</span></span> <span data-ttu-id="5cb9a-158">Por exemplo, se sua localidade padrão for `en-us`, não é possível especificar também uma substituição para `en-us`.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-158">For example, if your default locale is `en-us` you cannot not specify an  override for `en-us` as well.</span></span> 

<span data-ttu-id="5cb9a-p113">O exemplo a seguir se aplica a uma substituição de localidade para o elemento [Description]. Primeiro especifica a localidade padrão `en-us` e uma descrição em inglês e, em seguida, especifica uma política de [Override] com uma descrição francesa para a localidade `fr-fr`:</span><span class="sxs-lookup"><span data-stu-id="5cb9a-p113">The following example applies a locale override for the [Description] element. It first specifies a default locale of `en-us` and an English description, and then specifies an [Override] statement with a French description for the `fr-fr` locale:</span></span>

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

<span data-ttu-id="5cb9a-p114">Isso significa que o suplemento pressupõe a localidade `en-us` como padrão. Os usuários veriam a descrição em inglês no atributo `DefaultValue` para todas as localidades, a menos que a localidade do computador cliente fosse `fr-fr`, nesse caso, eles veriam a descrição em francês.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-p114">This means that the add-in assumes the `en-us` locale by default. Users would see the English description in the `DefaultValue` attribute for all locales unless the client computer's locale is `fr-fr`, in which case they would see the French description.</span></span>

<span data-ttu-id="5cb9a-p115">No exemplo a seguir, o suplemento especifica uma imagem separada mais apropriada para a localidade e a cultura `fr-fr`. Os usuários verão a imagem DefaultLogo.png por padrão, exceto quando a localidade do computador cliente for `fr-fr`. Nesse caso, os usuários veriam a imagem FrenchLogo.png.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-p115">In the following example, the add-in specifies a separate image that's more appropriate for the `fr-fr` locale and culture. Users see the image DefaultLogo.png by default, except when the locale of the client computer is `fr-fr`. In this case, users would see the image FrenchLogo.png.</span></span> 


```xml
<!-- Replace "domain" with a real web server name and path. -->
<IconUrl DefaultValue="https://<domain>/DefaultLogo.png"/>
<Override Locale="fr-fr" Value="https://<domain>/FrenchLogo.png"/>
```

<span data-ttu-id="5cb9a-p116">O exemplo a seguir mostra como localizar um recurso na seção `Resources`. Ele aplica um substituto local para uma imagem que é mais apropriada para a cultura `ja-jp`.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-p116">The following example shows how to localize a resource in the `Resources` section. It applies a locale override for an image that is more appropriate for the `ja-jp` culture.</span></span>

```xml
<Resources>
      <bt:Images>
        <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
        </bt:Image>
 ...
```


<span data-ttu-id="5cb9a-p117">Para o elemento [SourceLocation], o suporte a localidades adicionais significa fornecer um arquivo HTML de origem separado para cada um dos locais especificados. Os usuários de cada localidade que você especificar poderão ver uma página da Web personalizada que foi projetada para eles.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-p117">For the [SourceLocation] element, supporting additional locales means providing a separate source HTML file for each of the specified locales. Users in each locale you specify can see a customized webpage that you design for that them.</span></span>

<span data-ttu-id="5cb9a-p118">Para suplementos do Outlook, o elemento [SourceLocation] também atribui o fator forma, o que permite que você forneça um arquivo HTML de origem localizado e distinto para cada fator de foram correspondente. Você pode especificar um ou mais elementos filho [Override] em cada configuração aplicável ([DesktopSettings], [TabletSettings] ou [PhoneSettings]). O exemplo a seguir mostra os elementos de configurações para fatores de forma de desktop, tablet e smartphone, cada um com um arquivo HTML para a localidade padrão e outro para a localidade francesa.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-p118">For Outlook add-ins, the [SourceLocation] element also aligns to the form factor. This allows you to provide a separate, localized source HTML file for each corresponding form factor. You can specify one or more [Override] child elements in each applicable settings element ([DesktopSettings], [TabletSettings], or [PhoneSettings]). The following example shows settings elements for the desktop, tablet, and smartphone form factors, each with one HTML file for the default locale and another for the French locale.</span></span>


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

## <a name="localize-extended-overrides"></a><span data-ttu-id="5cb9a-174">Localize substituições estendidas</span><span class="sxs-lookup"><span data-stu-id="5cb9a-174">Localize extended overrides</span></span>

<span data-ttu-id="5cb9a-175">Alguns recursos de extensibilidade de Complementos do Office, como atalhos de teclado, são configurados com arquivos JSON hospedados em seu servidor, em vez de com o manifesto XML do complemento.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-175">Some extensibility features of Office Add-ins, such as keyboard shortcuts, are configured with JSON files that are hosted on your server, instead of with the add-in's XML manifest.</span></span> <span data-ttu-id="5cb9a-176">Esta seção pressupõe que você esteja familiarizado com substituições estendidas.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-176">This section assumes that you're familiar with extended overrides.</span></span> <span data-ttu-id="5cb9a-177">Consulte [Trabalhar com substituições estendidas do manifesto](extended-overrides.md) e do elemento [ExtendedOverrides.](../reference/manifest/extendedoverrides.md)</span><span class="sxs-lookup"><span data-stu-id="5cb9a-177">See [Work with extended overrides of the manifest](extended-overrides.md) and [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span>

<span data-ttu-id="5cb9a-178">Use o `ResourceUrl` atributo do [elemento ExtendedOverrides](../reference/manifest/extendedoverrides.md) para apontar o Office para um arquivo de recursos localizados.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-178">Use the `ResourceUrl` attribute of the [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element to point Office to a file of localized resources.</span></span> <span data-ttu-id="5cb9a-179">Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-179">The following is an example.</span></span>

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json" 
                       ResourceUrl="https://contoso.com/addin/my-resources.json">
    </ExtendedOverrides>
</OfficeApp>
```

<span data-ttu-id="5cb9a-180">O arquivo de substituições estendidas usa tokens em vez de cadeias de caracteres.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-180">The extended overrides file then uses tokens instead of strings.</span></span> <span data-ttu-id="5cb9a-181">As cadeias de caracteres de nomes de tokens no arquivo de recursos.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-181">The tokens name strings in the resource file.</span></span> <span data-ttu-id="5cb9a-182">Veja a seguir um exemplo que atribui um atalho de teclado a uma função (definida em outro lugar) que exibe o painel de tarefas do complemento.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-182">The following is an example that assigns a keyboard shortcut to a function (defined elsewhere) that displays the add-in's task pane.</span></span> <span data-ttu-id="5cb9a-183">Observação sobre essa marcação:</span><span class="sxs-lookup"><span data-stu-id="5cb9a-183">Note about this markup:</span></span>

- <span data-ttu-id="5cb9a-184">O exemplo não é muito válido.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-184">The example isn't quite valid.</span></span> <span data-ttu-id="5cb9a-185">(Adicionamos uma propriedade adicional necessária a ela abaixo.)</span><span class="sxs-lookup"><span data-stu-id="5cb9a-185">(We add a required additional property to it below.)</span></span>
- <span data-ttu-id="5cb9a-186">Os tokens devem ter o formato **${resource.*name-of-resource*}**.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-186">The tokens must have the format **${resource.*name-of-resource*}**.</span></span>

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

<span data-ttu-id="5cb9a-187">O arquivo de recursos, que também é formatado por JSON, tem uma propriedade de nível superior que é dividida em `resources` subpropropriedades por localidade.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-187">The resource file, which is also JSON-formatted, has a top-level `resources` property that is divided into subproperties by locale.</span></span> <span data-ttu-id="5cb9a-188">Para cada localidade, uma cadeia de caracteres é atribuída a cada token que foi usado no arquivo de substituições estendidas.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-188">For each locale, a string is assigned to each token that was used in the extended overrides file.</span></span> <span data-ttu-id="5cb9a-189">A seguir está um exemplo que tem cadeias de caracteres `en-us` para `fr-fr` e .</span><span class="sxs-lookup"><span data-stu-id="5cb9a-189">The following is an example which has strings for `en-us` and `fr-fr`.</span></span> <span data-ttu-id="5cb9a-190">Neste exemplo, o atalho de teclado é o mesmo em ambas as localidades, mas isso nem sempre será o caso, especialmente quando você estiver localizando para localidades que têm um alfabeto ou um sistema de escrita diferente e, portanto, um teclado diferente.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-190">In this example, the keyboard shortcut is the same in both locales, but that won't always be the case, especially when you are localizing for locales that have a different alphabet or writing system, and hence a different keyboard.</span></span>

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

<span data-ttu-id="5cb9a-191">Não há `default` nenhuma propriedade no arquivo que seja um ponto para as `en-us` `fr-fr` seções e.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-191">There is no `default` property in the file that is a peer to the `en-us` and `fr-fr` sections.</span></span> <span data-ttu-id="5cb9a-192">Isso acontece porque as cadeias de caracteres padrão, que são usadas quando a localidade do aplicativo host do Office não corresponder a nenhuma das propriedades *ll-cc* no arquivo de recursos, devem ser *definidas* no próprio arquivo de substituições estendido.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-192">This is because the default strings, which are used when the locale of the Office host application doesn't match any of the *ll-cc* properties in the resources file, *must be defined in the extended overrides file itself*.</span></span> <span data-ttu-id="5cb9a-193">Definir as cadeias de caracteres padrão diretamente no arquivo de substituições estendidas garante que o Office não baixe o arquivo de recurso quando a localidade do aplicativo do Office corresponde à localidade padrão do add-in (conforme especificado no manifesto).</span><span class="sxs-lookup"><span data-stu-id="5cb9a-193">Defining the default strings directly in the extended overrides file ensures that Office doesn't download the resource file when the locale of the Office application matches the default locale of the add-in (as specified in the manifest).</span></span> <span data-ttu-id="5cb9a-194">A seguir está uma versão corrigida do exemplo anterior de um arquivo de substituições estendido que usa tokens de recurso.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-194">The following is a corrected version of the preceding example of an extended overrides file that uses resource tokens.</span></span>

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

## <a name="match-datetime-format-with-client-locale"></a><span data-ttu-id="5cb9a-195">Fazer a correspondência entre o formato de data/hora e a localidade do cliente</span><span class="sxs-lookup"><span data-stu-id="5cb9a-195">Match date/time format with client locale</span></span>

<span data-ttu-id="5cb9a-196">Você pode obter a localidade da interface do usuário do aplicativo cliente do Office usando a **[propriedade displayLanguage.]**</span><span class="sxs-lookup"><span data-stu-id="5cb9a-196">You can get the locale of the user interface of the Office client application by using the **[displayLanguage]** property.</span></span> <span data-ttu-id="5cb9a-197">Em seguida, você pode exibir valores de data e hora em um formato consistente com a localidade atual do aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-197">You can then display date and time values in a format consistent with the current locale of the Office application.</span></span> <span data-ttu-id="5cb9a-198">Uma maneira de fazer isso é preparar um arquivo de recurso que especifica o formato de exibição de data/hora a ser usado em cada localidade com suporte do seu Suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-198">One way to do that is to prepare a resource file that specifies the date/time display format to use for each locale that your Office Add-in supports.</span></span> <span data-ttu-id="5cb9a-199">Em tempo de executar, o seu complemento pode usar o arquivo de recurso e corresponder ao formato de data/hora apropriado com a localidade obtida da **[propriedade displayLanguage.]**</span><span class="sxs-lookup"><span data-stu-id="5cb9a-199">At run time, your add-in can use the resource file and match the appropriate date/time format with the locale obtained from the **[displayLanguage]** property.</span></span>

<span data-ttu-id="5cb9a-200">Você pode obter a localidade dos dados do aplicativo cliente do Office usando a [propriedade contentLanguage.]</span><span class="sxs-lookup"><span data-stu-id="5cb9a-200">You can get the locale of the data of the Office client application by using the [contentLanguage] property.</span></span> <span data-ttu-id="5cb9a-201">Com base nesse valor, você pode, então, interpretar ou exibir adequadamente as cadeias de caracteres de data/hora.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-201">Based on this value, you can then appropriately interpret or display date/time strings.</span></span> <span data-ttu-id="5cb9a-202">Por exemplo, a localidade `jp-JP` expressa valores de data/hora como `yyyy/MM/dd`, e a localidade `fr-FR` como `dd/MM/yyyy`.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-202">For example, the `jp-JP` locale expresses data/time values as `yyyy/MM/dd`, and the `fr-FR` locale, `dd/MM/yyyy`.</span></span>


## <a name="use-ajax-for-globalization-and-localization"></a><span data-ttu-id="5cb9a-203">Usar o Ajax para a globalização e a localização</span><span class="sxs-lookup"><span data-stu-id="5cb9a-203">Use Ajax for globalization and localization</span></span>


<span data-ttu-id="5cb9a-204">Se você usar o Visual Studio para criar Suplementos do Office, o .NET Framework e Ajax fornecem maneiras de globalizar e localizar arquivos de script de cliente.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-204">If you use Visual Studio to create Office Add-ins, the .NET Framework and Ajax provide ways to globalize and localize client script files.</span></span>

<span data-ttu-id="5cb9a-p127">Você pode globalizar e utilizar as extensões do tipo JavaScript de [Data](/previous-versions/bb310850(v=vs.140)) e [Número](/previous-versions/bb310835(v=vs.140)) e o objeto [Data](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) do JavaScript no código do JavaScript para um suplemento do Office para exibir valores com base nas configurações de localização do navegador atual. Para saber mais, confira [Passo a passo: como globalizar uma data usando o script de cliente](/previous-versions/bb386581(v=vs.140)).</span><span class="sxs-lookup"><span data-stu-id="5cb9a-p127">You can globalize and use the [Date](/previous-versions/bb310850(v=vs.140)) and [Number](/previous-versions/bb310835(v=vs.140)) JavaScript type extensions and the JavaScript [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object in the JavaScript code for an Office Add-in to display values based on the locale settings on the current browser. For more information, see [Walkthrough: Globalizing a Date by Using Client Script](/previous-versions/bb386581(v=vs.140)).</span></span>

<span data-ttu-id="5cb9a-p128">Você pode incluir cadeias de caracteres de recurso localizadas diretamente em arquivos de JavaScript autônomos para fornecer arquivos de script de cliente para diferentes locais, que são definidos no navegador ou fornecidos pelo usuário. Crie um arquivo de script separado para cada localidade com suporte. Em cada arquivo de script, inclua um objeto no formato JSON que contenha as cadeias de caracteres de recursos para essa localidade. Os valores localizados serão aplicados quando o script for executado no navegador.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-p128">You can include localized resource strings directly in standalone JavaScript files to provide client script files for different locales, which are set on the browser or provided by the user. Create a separate script file for each supported locale. In each script file, include an object in JSON format that contains the resource strings for that locale. The localized values are applied when the script runs in the browser.</span></span>


## <a name="example-build-a-localized-office-add-in"></a><span data-ttu-id="5cb9a-211">Exemplo: Criar um Suplemento do Office localizado</span><span class="sxs-lookup"><span data-stu-id="5cb9a-211">Example: Build a localized Office Add-in</span></span>

<span data-ttu-id="5cb9a-212">Esta seção fornece exemplos que mostram como localizar uma descrição do Suplemento do Office, o nome de exibição e interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-212">This section provides examples that show you how to localize an Office Add-in description, display name, and UI.</span></span> 

> [!NOTE]
> <span data-ttu-id="5cb9a-213">Para baixar Visual Studio 2019, consulte a [página Visual Studio IDE .](https://visualstudio.microsoft.com/vs/)</span><span class="sxs-lookup"><span data-stu-id="5cb9a-213">To download Visual Studio 2019, see the [Visual Studio IDE page](https://visualstudio.microsoft.com/vs/).</span></span> <span data-ttu-id="5cb9a-214">Durante a instalação, você precisará selecionar a carga de trabalho de desenvolvimento do Office/SharePoint.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-214">During installation you'll need to select the Office/SharePoint development workload.</span></span>

### <a name="configure-office-to-use-additional-languages-for-display-or-editing"></a><span data-ttu-id="5cb9a-215">Configurar o Office para usar idiomas adicionais para exibição ou edição</span><span class="sxs-lookup"><span data-stu-id="5cb9a-215">Configure Office to use additional languages for display or editing</span></span>

<span data-ttu-id="5cb9a-216">Para executar o código de exemplo fornecido, configure o Office em seu computador para usar idiomas adicionais para que você possa testar seu complemento alternando o idioma usado para exibição em menus e comandos, para edição e revisão de texto, ou ambos.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-216">To run the sample code provided, configure Office on your computer to use additional languages so that you can test your add-in by switching the language used for display in menus and commands, for editing and proofing, or both.</span></span>

<span data-ttu-id="5cb9a-217">Você pode usar um Office Language Pack para instalar um idioma adicional.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-217">You can use an Office Language pack to install an additional language.</span></span> <span data-ttu-id="5cb9a-218">Para saber mais sobre os Pacotes de Idiomas e onde obtê-los, veja [Language Accessory Pack do Office](https://support.microsoft.com/office/82ee1236-0f9a-45ee-9c72-05b026ee809f).</span><span class="sxs-lookup"><span data-stu-id="5cb9a-218">For more information about Language Packs and where to get them, see [Language Accessory Pack for Office](https://support.microsoft.com/office/82ee1236-0f9a-45ee-9c72-05b026ee809f).</span></span>

<span data-ttu-id="5cb9a-219">Depois de instalar o Language Accessory Pack, você pode configurar o Office para usar o idioma instalado para exibir na interface do usuário, para edição de conteúdo do documento ou ambos.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-219">After you install the Language Accessory Pack, you can configure Office to use the installed language for display in the UI, for editing document content, or both.</span></span> <span data-ttu-id="5cb9a-220">O exemplo neste artigo usa uma instalação do Office que tenha o Pacote de Idiomas de espanhol aplicado.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-220">The example in this article uses an installation of Office that has the Spanish Language Pack applied.</span></span>

### <a name="create-an-office-add-in-project"></a><span data-ttu-id="5cb9a-221">Criar um projeto de Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="5cb9a-221">Create an Office Add-in project</span></span>

<span data-ttu-id="5cb9a-222">Você precisará criar um projeto de Visual Studio do Office 2019.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-222">You'll need to create a Visual Studio 2019 Office Add-in project.</span></span>

> [!NOTE]
> <span data-ttu-id="5cb9a-223">Se você ainda não instalou o Visual Studio 2019, consulte a página Visual Studio [IDE](https://visualstudio.microsoft.com/vs/) para obter instruções de download.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-223">If you haven't installed Visual Studio 2019, see the [Visual Studio IDE page](https://visualstudio.microsoft.com/vs/) for download instructions.</span></span> <span data-ttu-id="5cb9a-224">Durante a instalação, você precisará selecionar a carga de trabalho de desenvolvimento do Office/SharePoint.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-224">During installation you'll need to select the Office/SharePoint development workload.</span></span> <span data-ttu-id="5cb9a-225">Se você tiver instalado o Visual Studio 2019, use o Visual Studio [Instalador](/visualstudio/install/modify-visual-studio/) para garantir que a carga de trabalho de desenvolvimento do Office/SharePoint está instalada.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-225">If you have previously installed Visual Studio 2019, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio/) to ensure that the Office/SharePoint development workload is installed.</span></span>

1. <span data-ttu-id="5cb9a-226">Escolha **Criar um novo projeto**.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-226">Choose **Create a new project**.</span></span>

2. <span data-ttu-id="5cb9a-227">Usando a caixa de pesquisa, insira **suplemento**.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-227">Using the search box, enter **add-in**.</span></span> <span data-ttu-id="5cb9a-228">Escolha **Suplemento do Word Web**, em seguida, selecione **Próximo**.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-228">Choose **Word Web Add-in**, then select **Next**.</span></span>

3. <span data-ttu-id="5cb9a-229">Nomeia **seu projeto WorldReadyAddIn** e selecione **Criar**.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-229">Name your project **WorldReadyAddIn** and select **Create**.</span></span>

4. <span data-ttu-id="5cb9a-230">O Visual Studio cria uma solução, e os dois projetos dele são exibidos no **Gerenciador de Soluções**.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-230">Visual Studio creates a solution and its two projects appear in **Solution Explorer**.</span></span> <span data-ttu-id="5cb9a-231">O arquivo **Home.html** é aberto no Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-231">The **Home.html** file opens in Visual Studio.</span></span>


### <a name="localize-the-text-used-in-your-add-in"></a><span data-ttu-id="5cb9a-232">Localizar o texto usado no seu suplemento</span><span class="sxs-lookup"><span data-stu-id="5cb9a-232">Localize the text used in your add-in</span></span>

<span data-ttu-id="5cb9a-233">O texto que você deseja localizar para outro idioma aparece em duas áreas:</span><span class="sxs-lookup"><span data-stu-id="5cb9a-233">The text that you want to localize for another language appears in two areas:</span></span>

-  <span data-ttu-id="5cb9a-p135">**Nome de exibição e descrição do suplemento**. São controlados por entradas no arquivo de manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-p135">**Add-in display name and description**. This is controlled by entries in the add-in manifest file.</span></span>

-  <span data-ttu-id="5cb9a-236">**Interface do usuário do suplemento**.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-236">**Add-in UI**.</span></span> <span data-ttu-id="5cb9a-237">Você pode localizar as cadeias de caracteres que aparecem na interface do usuário do seu suplemento usando códigos do JavaScript, por exemplo, usando um arquivo de recurso separado que contenha as cadeias de caracteres localizadas.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-237">You can localize the strings that appear in your add-in UI by using JavaScript code, for example, by using a separate resource file that contains the localized strings.</span></span>

<span data-ttu-id="5cb9a-238">Para localizar o nome de exibição e a descrição do suplemento:</span><span class="sxs-lookup"><span data-stu-id="5cb9a-238">To localize the add-in display name and description:</span></span>

1. <span data-ttu-id="5cb9a-239">No **Explorador de** Soluções, **expanda WorldReadyAddIn,** **WorldReadyAddInManifest** e escolha **WorldReadyAddIn.xml**.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-239">In **Solution Explorer**, expand **WorldReadyAddIn**, **WorldReadyAddInManifest**, and then choose **WorldReadyAddIn.xml**.</span></span>

2. <span data-ttu-id="5cb9a-240">No arquivo WorldReadyAddInManifest.xml, substitua os elementos [DisplayName] e [Description] pelo seguinte bloco de código:</span><span class="sxs-lookup"><span data-stu-id="5cb9a-240">In WorldReadyAddInManifest.xml, replace the [DisplayName] and [Description] elements with the following block of code:</span></span>

    > [!NOTE]
    > <span data-ttu-id="5cb9a-241">Você pode substituir as cadeias de caracteres do idioma espanhol localizado usadas neste exemplo dos elementos [DisplayName] e [Description] pelas cadeias de caracteres localizados para qualquer outro idioma.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-241">You can replace the Spanish language localized strings used in this example for the [DisplayName] and [Description] elements with the localized strings for any other language.</span></span>

    ```xml
    <DisplayName DefaultValue="World Ready add-in">
      <Override Locale="es-es" Value="Aplicación de uso internacional"/>
    </DisplayName>
    <Description DefaultValue="An add-in for testing localization">
      <Override Locale="es-es" Value="Una aplicación para la prueba de la localización"/>
    </Description>
    ```

3. <span data-ttu-id="5cb9a-242">Quando você altera o idioma de exibição do Office 2013 do inglês para o espanhol, por exemplo, e executa o suplemento, o nome de exibição do suplemento e a descrição são mostrados com texto localizado.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-242">When you change the display language for Office 2013 from English to Spanish, for example, and then run the add-in, the add-in display name and description are shown with localized text.</span></span>

<span data-ttu-id="5cb9a-243">Para definir a interface do usuário do suplemento:</span><span class="sxs-lookup"><span data-stu-id="5cb9a-243">To lay out the add-in UI:</span></span>

1. <span data-ttu-id="5cb9a-244">No Visual Studio, no **Gerenciador de Soluções**, selecione **Home.html**.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-244">In Visual Studio, in **Solution Explorer**, choose **Home.html**.</span></span>

2. <span data-ttu-id="5cb9a-245">Substitua o conteúdo do elemento `<body>` no Home.html com o HTML a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-245">Replace the `<body>` element contents in Home.html with the following HTML, and save the file.</span></span>

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

<span data-ttu-id="5cb9a-246">A figura a seguir mostra o elemento do cabeçalho (h1) e o elemento do parágrafo (p) que exibirá o texto localizado quando concluir as etapas restantes e executar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-246">The following figure shows the heading (h1) element and the paragraph (p) element that will display localized text when you complete the remaining steps and run the add-in.</span></span>

<span data-ttu-id="5cb9a-247">*Figura 1. A interface do usuário do suplemento*</span><span class="sxs-lookup"><span data-stu-id="5cb9a-247">*Figure 1. The add-in UI*</span></span>

![Interface de usuário do aplicativo com as seções realçadas.](../images/office15-app-how-to-localize-fig03.png)

### <a name="add-the-resource-file-that-contains-the-localized-strings"></a><span data-ttu-id="5cb9a-249">Adicionar o arquivo de recurso que contém as cadeias de caracteres localizadas</span><span class="sxs-lookup"><span data-stu-id="5cb9a-249">Add the resource file that contains the localized strings</span></span>

<span data-ttu-id="5cb9a-250">O arquivo de recurso do JavaScript contém as cadeias de caracteres usadas para a interface do usuário do suplemento.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-250">The JavaScript resource file contains the strings used for the add-in UI.</span></span> <span data-ttu-id="5cb9a-251">O HTML da interface do usuário do suplemento de amostra tem um elemento `<h1>` que exibe uma saudação e um elemento `<p>` que apresenta o suplemento ao usuário.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-251">The HTML for the sample add-in UI contains an `<h1>` element that displays a greeting, and a `<p>` element that introduces the add-in to the user.</span></span> 

<span data-ttu-id="5cb9a-p138">Para habilitar cadeias de caracteres para o cabeçalho e parágrafo, coloque as cadeias de caracteres em um arquivo de recurso separado. O arquivo de recurso cria um objeto do JavaScript que contém um objeto JSON (JavaScript Object Notation) separado para cada conjunto de cadeias de caracteres localizadas. O arquivo de recurso também fornece um método para obter o objeto JSON apropriado de volta para uma determinada localidade.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-p138">To enable localized strings for the heading and paragraph, you place the strings in a separate resource file. The resource file creates a JavaScript object that contains a separate JavaScript Object Notation (JSON) object for each set of localized strings. The resource file also provides a method for getting back the appropriate JSON object for a given locale.</span></span>

<span data-ttu-id="5cb9a-255">Para adicionar o arquivo de recurso ao projeto do suplemento:</span><span class="sxs-lookup"><span data-stu-id="5cb9a-255">To add the resource file to the add-in project:</span></span>

1. <span data-ttu-id="5cb9a-256">No **Gerenciador de Soluções** no Visual Studio, clique com o botão direito no projeto **WorldReadyAddInWeb** e escolha **Adicionar** > **Novo Item**.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-256">In **Solution Explorer** in Visual Studio, right-click the **WorldReadyAddInWeb** project and choose **Add** > **New Item**.</span></span> 

2. <span data-ttu-id="5cb9a-257">Na caixa de diálogo **Adicionar Novo Item**, escolha **Arquivo JavaScript**.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-257">In the **Add New Item** dialog box, choose **JavaScript File**.</span></span>

3. <span data-ttu-id="5cb9a-258">Insira **UIStrings.js** como nome do arquivo e escolha **Adicionar**.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-258">Enter **UIStrings.js** as the file name and choose **Add**.</span></span>

4. <span data-ttu-id="5cb9a-259">Adicione o código a seguir ao arquivo UIStrings.js e salve-o.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-259">Add the following code to the UIStrings.js file, and save the file.</span></span>

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

<span data-ttu-id="5cb9a-260">O arquivo de recurso UIStrings.js cria o objeto, **UIStrings**, que contém as cadeias de caracteres localizadas para a interface do usuário do suplemento.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-260">The UIStrings.js resource file creates an object, **UIStrings**, which contains the localized strings for your add-in UI.</span></span>

### <a name="localize-the-text-used-for-the-add-in-ui"></a><span data-ttu-id="5cb9a-261">Localizar o texto usado na interface do usuário do suplemento</span><span class="sxs-lookup"><span data-stu-id="5cb9a-261">Localize the text used for the add-in UI</span></span>

<span data-ttu-id="5cb9a-p139">Para usar o arquivo de recurso no seu suplemento, você precisará adicionar a ele uma marca de script em Home.html. Quando Home.html for carregado, o UIStrings.js será executado e o objeto **UIStrings** que você utiliza para obter a cadeia de caracteres ficará disponível para seu código. Adicione o seguinte HTML à marca de cabeçalho do Home.html para tornar **UIStrings** disponível para seu código.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-p139">To use the resource file in your add-in, you'll need to add a script tag for it on Home.html. When Home.html is loaded, UIStrings.js executes and the **UIStrings** object that you use to get the strings is available to your code. Add the following HTML in the head tag for Home.html to make **UIStrings** available to your code.</span></span>

```html
<!-- Resource file for localized strings: -->
<script src="../UIStrings.js" type="text/javascript"></script>
```

<span data-ttu-id="5cb9a-265">Agora você pode usar o objeto **UIStrings** para definir as cadeias de caracteres da interface do usuário do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-265">Now you can use the **UIStrings** object to set the strings for the UI of your add-in.</span></span>

<span data-ttu-id="5cb9a-266">Se você quiser alterar a localização do seu complemento com base em qual idioma é usado para exibição em menus e comandos no aplicativo cliente do Office, use a propriedade **Office.context.displayLanguage** para obter a localidade desse idioma.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-266">If you want to change the localization for your add-in based on what language is used for display in menus and commands in the Office client application, you use the **Office.context.displayLanguage** property to get the locale for that language.</span></span> <span data-ttu-id="5cb9a-267">Por exemplo, se o idioma do aplicativo usar espanhol para exibição em menus e comandos, a propriedade **Office.context.displayLanguage** retornará o código de idioma es-ES.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-267">For example, if the application language uses Spanish for display in menus and commands, the **Office.context.displayLanguage** property will return the language code es-ES.</span></span>

<span data-ttu-id="5cb9a-268">Se você quiser alterar a localização do seu complemento com base em qual idioma está sendo usado para editar conteúdo de documento, use a propriedade **Office.context.contentLanguage** para obter a localidade desse idioma.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-268">If you want to change the localization for your add-in based on what language is being used for editing document content, you use the **Office.context.contentLanguage** property to get the locale for that language.</span></span> <span data-ttu-id="5cb9a-269">Por exemplo, se o idioma do aplicativo usar espanhol para edição de conteúdo de documento, a propriedade **Office.context.contentLanguage** retornará o código de idioma es-ES.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-269">For example, if the application language uses Spanish for editing document content, the **Office.context.contentLanguage** property will return the language code es-ES.</span></span>

<span data-ttu-id="5cb9a-270">Depois de saber o idioma que o aplicativo está usando, você pode usar **UIStrings** para obter o conjunto de cadeias de caracteres localizadas que corresponde ao idioma do aplicativo.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-270">After you know the language the application is using, you can use **UIStrings** to get the set of localized strings that matches the application language.</span></span>

<span data-ttu-id="5cb9a-271">Substitua o código no arquivo Home.js pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-271">Replace the code in the Home.js file with the following code.</span></span> <span data-ttu-id="5cb9a-272">O código mostra como você pode alterar as cadeias de caracteres usadas nos elementos da interface do usuário no Home.html com base no idioma de exibição do aplicativo ou no idioma de edição do aplicativo.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-272">The code shows how you can change the strings used in the UI elements on Home.html based on either the display language of the application or the editing language of the application.</span></span>

> [!NOTE]
> <span data-ttu-id="5cb9a-273">Para alternar entre a alteração da localização do suplemento com base no idioma usado para edição, remova o comentário da linha de código `var myLanguage = Office.context.contentLanguage;` e inclua o comentário na linha de código `var myLanguage = Office.context.displayLanguage;`</span><span class="sxs-lookup"><span data-stu-id="5cb9a-273">To switch between changing the localization of the add-in based on the language used for editing, uncomment the line of code  `var myLanguage = Office.context.contentLanguage;` and comment out the line of code `var myLanguage = Office.context.displayLanguage;`</span></span>

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

### <a name="test-your-localized-add-in"></a><span data-ttu-id="5cb9a-274">Testar seu suplemento localizado</span><span class="sxs-lookup"><span data-stu-id="5cb9a-274">Test your localized add-in</span></span>

<span data-ttu-id="5cb9a-275">Para testar seu complemento localizado, altere o idioma usado para exibição ou edição no aplicativo do Office e execute o seu complemento.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-275">To test your localized add-in, change the language used for display or editing in the Office application and then run your add-in.</span></span>

<span data-ttu-id="5cb9a-276">Para alterar o idioma usado para exibir ou editar no seu suplemento:</span><span class="sxs-lookup"><span data-stu-id="5cb9a-276">To change the language used for display or editing in your add-in:</span></span>

1. <span data-ttu-id="5cb9a-277">No Word, escolha **Arquivo** > **Opções** > **Idioma**.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-277">In Word, choose **File** > **Options** > **Language**.</span></span> <span data-ttu-id="5cb9a-278">A figura a seguir mostra a caixa de diálogo **Opções do Word** aberta na guia Idioma.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-278">The following figure shows the **Word Options** dialog box opened to the Language tab.</span></span>

    <span data-ttu-id="5cb9a-279">*Figura 2. Opções de idioma na caixa de diálogo Opções do Word*</span><span class="sxs-lookup"><span data-stu-id="5cb9a-279">*Figure 2. Language options in the Word Options dialog box*</span></span>

    ![Caixa de diálogo Opções do Word](../images/office15-app-how-to-localize-fig04.png)

2. <span data-ttu-id="5cb9a-281">Em **Escolher Idioma de Exibição**, selecione o idioma desejado para exibição, por exemplo, espanhol, e selecione a seta para cima para mover o idioma espanhol para a primeira posição na lista.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-281">Under **Choose Display Language**, select the language that you want for display, for example Spanish, and then choose the up arrow to move the Spanish language to the first position in the list.</span></span> <span data-ttu-id="5cb9a-282">Como alternativa, para alterar o idioma usado para edição, em **Escolher** Idiomas de Edição, escolha o idioma que você deseja usar para edição, por exemplo, espanhol e escolha **Definir como Padrão**.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-282">Alternatively, to change the language used for editing, under **Choose Editing Languages**, choose the language you want to use for editing, for example, Spanish, and then choose **Set as Default**.</span></span>

3. <span data-ttu-id="5cb9a-283">Escolha **OK** para confirmar sua seleção e feche o Word.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-283">Choose **OK** to confirm your selection, and then close Word.</span></span>

4. <span data-ttu-id="5cb9a-284">Pressione **F5** no Visual Studio para executar o suplemento de amostra ou escolha **Depurar** > **Iniciar Depuração** na barra de menus.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-284">Press **F5** in Visual Studio to run the sample add-in, or choose **Debug** > **Start Debugging** from the menu bar.</span></span>

5. <span data-ttu-id="5cb9a-285">No Word, escolha **Página Inicial** > **Mostrar Painel de Tarefas**.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-285">In Word, choose **Home** > **Show Taskpane**.</span></span>

<span data-ttu-id="5cb9a-286">Depois de executar, as cadeias de caracteres na interface do usuário do complemento mudam para corresponder ao idioma usado pelo aplicativo, conforme mostrado na figura a seguir.</span><span class="sxs-lookup"><span data-stu-id="5cb9a-286">Once running, the strings in the add-in UI change to match the language used by the application, as shown in the following figure.</span></span>


<span data-ttu-id="5cb9a-287">*Figura 3. Interface do usuário do suplemento com o texto localizado*</span><span class="sxs-lookup"><span data-stu-id="5cb9a-287">*Figure 3. Add-in UI with localized text*</span></span>

![Aplicativo com texto localizado da interface do usuário.](../images/office15-app-how-to-localize-fig05.png)

## <a name="see-also"></a><span data-ttu-id="5cb9a-289">Confira também</span><span class="sxs-lookup"><span data-stu-id="5cb9a-289">See also</span></span>

- [<span data-ttu-id="5cb9a-290">Diretrizes de design para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="5cb9a-290">Design guidelines for Office Add-ins</span></span>](../design/add-in-design.md)
- <span data-ttu-id="5cb9a-291">[Identificadores de idioma e valores da ID de OptionState no Office 2013](/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15))</span><span class="sxs-lookup"><span data-stu-id="5cb9a-291">[Language identifiers and OptionState Id values in Office 2013](/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15))</span></span>

[DefaultLocale]:         ../reference/manifest/defaultlocale.md
[Descrição]:           ../reference/manifest/description.md
[Description]:           ../reference/manifest/description.md
[DisplayName]:           ../reference/manifest/displayname.md
[IconUrl]:               ../reference/manifest/iconurl.md
[HighResolutionIconUrl]: ../reference/manifest/highresolutioniconurl.md
[Resources]:             ../reference/manifest/resources.md
[SourceLocation]:        ../reference/manifest/sourcelocation.md
[Override]:              ../reference/manifest/override.md
[DesktopSettings]:       ../reference/manifest/desktopsettings.md
[TabletSettings]:        ../reference/manifest/tabletsettings.md
[PhoneSettings]:         ../reference/manifest/phonesettings.md
[displayLanguage]:       /javascript/api/office/office.context#displaylanguage
[contentLanguage]:       /javascript/api/office/office.context#contentlanguage
[RFC 3066]:              https://www.rfc-editor.org/info/rfc3066
