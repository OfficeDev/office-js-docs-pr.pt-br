---
title: Localização para Suplementos do Office
description: Você pode usar a API JavaScript para Office para determinar uma localidade e exibir cadeias de caracteres com base na localidade do aplicativo host ou para interpretar ou exibir dados com base na localidade dos dados.
ms.date: 10/11/2019
localization_priority: Normal
ms.openlocfilehash: c2404177f2188a505522d972d5bdfdf323394eba
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/22/2019
ms.locfileid: "37626765"
---
# <a name="localization-for-office-add-ins"></a><span data-ttu-id="2f827-103">Localização para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="2f827-103">Localization for Office Add-ins</span></span>

<span data-ttu-id="2f827-p101">Você pode implementar qualquer esquema de localização que seja apropriado para o seu Suplemento do Office. A API JavaScript e o esquema do manifesto da plataforma de Suplementos do Office oferecem algumas opções. Você pode usar a API JavaScript para Office para determinar uma localidade e exibir cadeias de caracteres com base na localidade do aplicativo host ou para interpretar ou exibir dados com base na localidade dos dados. Você pode usar o manifesto para especificar informações descritivas e o local do arquivo do suplemento específico da localidade. Como alternativa, você pode usar o script do Microsoft Ajax para dar suporte à globalização e localização.</span><span class="sxs-lookup"><span data-stu-id="2f827-p101">You can implement any localization scheme that's appropriate for your Office Add-in. The JavaScript API and manifest schema of the Office Add-ins platform provide some choices. You can use the JavaScript API for Office to determine a locale and display strings based on the locale of the host application, or to interpret or display data based on the locale of the data. You can use the manifest to specify locale-specific add-in file location and descriptive information. Alternatively, you can use Microsoft Ajax script to support globalization and localization.</span></span>

## <a name="use-the-javascript-api-to-determine-locale-specific-strings"></a><span data-ttu-id="2f827-109">Usar a API JavaScript para determinar cadeias de caracteres específicas da localidade</span><span class="sxs-lookup"><span data-stu-id="2f827-109">Use the JavaScript API to determine locale-specific strings</span></span>

<span data-ttu-id="2f827-110">A API JavaScript para Office fornece duas propriedades que oferecem suporte à exibição ou interpretação de valores consistentes com a localidade do aplicativo host e dos dados:</span><span class="sxs-lookup"><span data-stu-id="2f827-110">The JavaScript API for Office provides two properties that support displaying or interpreting values consistent with the locale of the host application and data:</span></span>

- <span data-ttu-id="2f827-111">[Context.displayLanguage][displayLanguage] especifica a localidade (ou o idioma) da interface do usuário do aplicativo host.</span><span class="sxs-lookup"><span data-stu-id="2f827-111">[Context.displayLanguage][displayLanguage] specifies the locale (or language) of the user interface of the host application.</span></span> <span data-ttu-id="2f827-112">O exemplo a seguir verifica se o aplicativo host usa a localidade en-US ou fr-FR e exibe uma saudação específica para a localidade.</span><span class="sxs-lookup"><span data-stu-id="2f827-112">The following example verifies if the host application uses the en-US or fr-FR locale, and displays a locale-specific greeting.</span></span>

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

- <span data-ttu-id="2f827-p103">[Context.contentLanguage][contentLanguage] especifica a localidade (ou o idioma) dos dados. Estendendo o último exemplo de código, em vez de verificar a propriedade [displayLanguage], atribua a `myLanguage` o valor da propriedade [contentLanguage] e use o restante do mesmo código para exibir uma saudação com base na localidade dos dados:</span><span class="sxs-lookup"><span data-stu-id="2f827-p103">[Context.contentLanguage][contentLanguage] specifies the locale (or language) of the data. Extending the last code sample, instead of checking the [displayLanguage] property, assign `myLanguage` the value of the [contentLanguage] property, and use the rest of the same code to display a greeting based on the locale of the data:</span></span>

    ```js
    var myLanguage = Office.context.contentLanguage;
    ```

## <a name="control-localization-from-the-manifest"></a><span data-ttu-id="2f827-115">Controlar a localização do manifesto</span><span class="sxs-lookup"><span data-stu-id="2f827-115">Control localization from the manifest</span></span>


<span data-ttu-id="2f827-p104">Cada Suplemento do Office especifica um elemento [DefaultLocale] e uma localidade em seu manifesto. Por padrão, a plataforma do Suplemento do Office e os aplicativos host do Office aplicam os valores dos elementos [Description], [DisplayName], [IconUrl], [HighResolutionIconUrl] e [SourceLocation] a todas as localidades. Como opção, você pode dar suporte a valores específicos para localidades específicas, especificando um elemento-filho [Override]para cada localidade adicional, para qualquer um desses cinco elementos. O valor do elemento [DefaultLocale] e do atributo `Locale` do elemento [Override] é especificado de acordo com o [RFC 3066], "Marcas para a Identificação dos Idiomas". A Tabela 1 descreve o suporte de localização para esses elementos.</span><span class="sxs-lookup"><span data-stu-id="2f827-p104">Every Office Add-in specifies a [DefaultLocale] element and a locale in its manifest. By default, the Office Add-in platform and Office host applications apply the values of the [Description], [DisplayName], [IconUrl], [HighResolutionIconUrl], and [SourceLocation] elements to all locales. You can optionally support specific values for specific locales, by specifying an [Override] child element for each additional locale, for any of these five elements. The value for the [DefaultLocale] element and for the `Locale` attribute of the [Override] element is specified according to [RFC 3066], "Tags for the Identification of Languages." Table 1 describes the localizing support for these elements.</span></span>

<span data-ttu-id="2f827-121">*Tabela 1. Suporte de localização*</span><span class="sxs-lookup"><span data-stu-id="2f827-121">*Table 1. Localization support*</span></span>


|<span data-ttu-id="2f827-122">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="2f827-122">**Element**</span></span>|<span data-ttu-id="2f827-123">**Suporte de localização**</span><span class="sxs-lookup"><span data-stu-id="2f827-123">**Localization support**</span></span>|
|:-----|:-----|
|<span data-ttu-id="2f827-124">[Descrição]</span><span class="sxs-lookup"><span data-stu-id="2f827-124">[Description]</span></span>   |<span data-ttu-id="2f827-125">Os usuários de cada localidade especificada podem ver uma descrição localizada do suplemento no AppSource (ou no catálogo privado).</span><span class="sxs-lookup"><span data-stu-id="2f827-125">Users in each locale you specify can see a localized description for the add-in in AppSource (or private catalog).</span></span><br/><span data-ttu-id="2f827-126">Para os suplementos do Outlook, os usuários podem ver a descrição no Centro de Administração do Exchange (EAC) após a instalação.</span><span class="sxs-lookup"><span data-stu-id="2f827-126">For Outlook add-ins, users can see the description in the Exchange Admin Center (EAC) after installation.</span></span>|
|<span data-ttu-id="2f827-127">[DisplayName]</span><span class="sxs-lookup"><span data-stu-id="2f827-127">[DisplayName]</span></span>   |<span data-ttu-id="2f827-128">Os usuários de cada localidade especificada podem ver uma descrição localizada do suplemento no AppSource (ou no catálogo privado).</span><span class="sxs-lookup"><span data-stu-id="2f827-128">Users in each locale you specify can see a localized description for the add-in in AppSource (or private catalog).</span></span><br/><span data-ttu-id="2f827-129">Para os suplementos do Outlook, os usuários podem ver o nome de exibição como um rótulo para o botão de suplemento do Outlook e no EAC após a instalação.</span><span class="sxs-lookup"><span data-stu-id="2f827-129">For Outlook add-ins, users can see the display name as a label for the Outlook add-in button and in the EAC after installation.</span></span><br/><span data-ttu-id="2f827-130">Para os suplementos do painel de tarefas e do conteúdo, os usuários podem ver o nome de exibição na faixa de opções após a instalação do suplemento.</span><span class="sxs-lookup"><span data-stu-id="2f827-130">For content and task pane add-ins, users can see the display name in the ribbon after installing the add-in.</span></span>|
|<span data-ttu-id="2f827-131">[IconUrl]</span><span class="sxs-lookup"><span data-stu-id="2f827-131">[IconUrl]</span></span>        |<span data-ttu-id="2f827-p105">A imagem do ícone é opcional. Você pode usar a mesma técnica de substituição para especificar uma determinada imagem para uma cultura específica. Se você usar e localizar um ícone, os usuários em cada localidade que você especificar poderão ver uma imagem de ícone localizada para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="2f827-p105">The icon image is optional. You can use the same override technique to specify a certain image for a specific culture. If you use and localize an icon, users in each locale you specify can see a localized icon image for the add-in.</span></span><br/><span data-ttu-id="2f827-135">Para suplementos do Outlook, os usuários podem ver o ícone no EAC depois de instalar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="2f827-135">For Outlook add-ins, users can see the icon in the EAC after installing the add-in.</span></span><br/><span data-ttu-id="2f827-136">Para os suplementos do painel de tarefas e do conteúdo, os usuários podem ver o ícone na faixa de opções após a instalação do suplemento.</span><span class="sxs-lookup"><span data-stu-id="2f827-136">For content and task pane add-ins, users can see the icon in the ribbon after installing the add-in.</span></span>|
|<span data-ttu-id="2f827-137">[HighResolutionIconUrl] **Importante:** este elemento só fica disponível ao usar a versão 1.1 do manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="2f827-137">[HighResolutionIconUrl] **Important:** This element is available only when using add-in manifest version 1.1.</span></span>|<span data-ttu-id="2f827-p106">A imagem do ícone de alta resolução é opcional, mas se ela for especificada, deverá ocorrer após o elemento [IconUrl]. Quando [HighResolutionIconUrl] for especificado e o suplemento estiver instalado em um dispositivo que ofereça suporte à resolução dpi alto, o valor [HighResolutionIconUrl] é usado em vez do valor para [IconUrl].</span><span class="sxs-lookup"><span data-stu-id="2f827-p106">The high resolution icon image is optional but if it is specified, it must occur after the  [IconUrl] element. When [HighResolutionIconUrl] is specified, and the add-in is installed on a device that supports high dpi resolution, the [HighResolutionIconUrl] value is used instead of the value for [IconUrl].</span></span><br/><span data-ttu-id="2f827-p107">Você pode usar a mesma técnica de substituição para especificar uma determinada imagem para uma cultura específica. Se você usar e localizar um ícone, os usuários em cada localidade que você especificar podem ver uma imagem de ícone localizada para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="2f827-p107">You can use the same override technique to specify a certain image for a specific culture. If you use and localize an icon, users in each locale you specify can see a localized icon image for the add-in.</span></span><br/><span data-ttu-id="2f827-142">Para suplementos do Outlook, os usuários podem ver o ícone no EAC depois de instalar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="2f827-142">For Outlook add-ins, users can see the icon in the EAC after installing the add-in.</span></span><br/><span data-ttu-id="2f827-143">Para os suplementos do painel de tarefas e do conteúdo, os usuários podem ver o ícone na faixa de opções após a instalação do suplemento.</span><span class="sxs-lookup"><span data-stu-id="2f827-143">For content and task pane add-ins, users can see the icon in the ribbon after installing the add-in.</span></span>|
|<span data-ttu-id="2f827-144">[Recursos] **Importante:** este elemento só fica disponível ao usar a versão 1.1 do manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="2f827-144">[Resources] **Important:** This element is available only when using add-in manifest version 1.1.</span></span>   |<span data-ttu-id="2f827-145">Os usuários em cada localidade especificada podem ver recursos de cadeias de caracteres e de ícones que você projetou especificamente para o suplemento dessa localidade.</span><span class="sxs-lookup"><span data-stu-id="2f827-145">Users in each locale you specify can see string and icon resources that you specifically create for the add-in for that locale.</span></span> |
|<span data-ttu-id="2f827-146">[SourceLocation]</span><span class="sxs-lookup"><span data-stu-id="2f827-146">[SourceLocation]</span></span>   |<span data-ttu-id="2f827-147">Os usuários de cada localidade especificada podem ver a página da Web que você projetou especificamente para o suplemento dessa localidade.</span><span class="sxs-lookup"><span data-stu-id="2f827-147">Users in each locale you specify can see a webpage that you specifically design for the add-in for that locale.</span></span> |


> [!NOTE]
> <span data-ttu-id="2f827-148">Você só pode localizar o nome de exibição e a descrição das localidades para as quais o Office oferece suporte.</span><span class="sxs-lookup"><span data-stu-id="2f827-148">You can localize the description and display name for only the locales that Office supports.</span></span> <span data-ttu-id="2f827-149">Consulte [Identificadores de idioma e valores de OptionState Id no Office 2013](/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15)) para obter uma lista de idiomas e localidades da versão atual do Office.</span><span class="sxs-lookup"><span data-stu-id="2f827-149">See [Language identifiers and OptionState Id values in Office 2013](/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15)) for a list of languages and locales for the current release of Office.</span></span>


### <a name="examples"></a><span data-ttu-id="2f827-150">Exemplos</span><span class="sxs-lookup"><span data-stu-id="2f827-150">Examples</span></span>

<span data-ttu-id="2f827-151">Por exemplo, um Suplemento do Office pode especificar [DefaultLocale] como `en-us`.</span><span class="sxs-lookup"><span data-stu-id="2f827-151">For example, an Office Add-in can specify the [DefaultLocale] as `en-us`.</span></span> <span data-ttu-id="2f827-152">Para o elemento [DisplayName], o suplemento pode especificar um elemento filho [Override] para a localidade `fr-fr`, como mostrado abaixo.</span><span class="sxs-lookup"><span data-stu-id="2f827-152">For the [DisplayName] element, the add-in can specify an [Override] child element for the locale `fr-fr`, as shown below.</span></span>


```xml
<DefaultLocale>en-us</DefaultLocale>
...
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

> [!NOTE]
> <span data-ttu-id="2f827-153">Se for preciso localizar para mais de uma área dentro de uma família de idiomas, como `de-de` e `de-at`, recomendamos que você use elementos `Override` separados para cada área.</span><span class="sxs-lookup"><span data-stu-id="2f827-153">If you need to localize for more than one area within a language family, such as `de-de` and `de-at`, we recommend that you use separate `Override` elements for each area.</span></span> <span data-ttu-id="2f827-154">Usar apenas o nome do idioma, nesse caso, `de`, não tem suporte em todas as combinações de plataformas e aplicativos host do Office.</span><span class="sxs-lookup"><span data-stu-id="2f827-154">Using just the language name alone, in this case, `de`, is not supported across all combinations of Office host applications and platforms.</span></span>

<span data-ttu-id="2f827-p111">Isso significa que o suplemento pressupõe a localidade `en-us` como padrão. Os usuários veem o nome de exibição em inglês "Video player" para todas as localidades, a menos que a localidade do computador cliente fosse `fr-fr`, nesse caso os usuários veria o nome de exibição em francês "Lecteur vidéo".</span><span class="sxs-lookup"><span data-stu-id="2f827-p111">This means that the add-in assumes the  `en-us` locale by default. Users see the English display name of "Video player" for all locales unless the client computer's locale is `fr-fr`, in which case users would see the French display name "Lecteur vidéo".</span></span>

> [!NOTE]
> <span data-ttu-id="2f827-157">Você só pode especificar uma única substituição por idioma, inclusive para a localidade padrão.</span><span class="sxs-lookup"><span data-stu-id="2f827-157">You may only specify a single override per language, including for the default locale.</span></span> <span data-ttu-id="2f827-158">Por exemplo, se sua localidade padrão for `en-us`, não é possível especificar também uma substituição para `en-us`.</span><span class="sxs-lookup"><span data-stu-id="2f827-158">For example, if your default locale is `en-us` you cannot not specify an  override for `en-us` as well.</span></span> 

<span data-ttu-id="2f827-p113">O exemplo a seguir se aplica a uma substituição de localidade para o elemento [Description]. Primeiro especifica a localidade padrão `en-us` e uma descrição em inglês e, em seguida, especifica uma política de [Override] com uma descrição francesa para a localidade `fr-fr`:</span><span class="sxs-lookup"><span data-stu-id="2f827-p113">The following example applies a locale override for the [Description] element. It first specifies a default locale of `en-us` and an English description, and then specifies an [Override] statement with a French description for the `fr-fr` locale:</span></span>

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

<span data-ttu-id="2f827-p114">Isso significa que o suplemento pressupõe a localidade `en-us` como padrão. Os usuários veriam a descrição em inglês no atributo `DefaultValue` para todas as localidades, a menos que a localidade do computador cliente fosse `fr-fr`, nesse caso, eles veriam a descrição em francês.</span><span class="sxs-lookup"><span data-stu-id="2f827-p114">This means that the add-in assumes the `en-us` locale by default. Users would see the English description in the `DefaultValue` attribute for all locales unless the client computer's locale is `fr-fr`, in which case they would see the French description.</span></span>

<span data-ttu-id="2f827-p115">No exemplo a seguir, o suplemento especifica uma imagem separada mais apropriada para a localidade e a cultura `fr-fr`. Os usuários verão a imagem DefaultLogo.png por padrão, exceto quando a localidade do computador cliente for `fr-fr`. Nesse caso, os usuários veriam a imagem FrenchLogo.png.</span><span class="sxs-lookup"><span data-stu-id="2f827-p115">In the following example, the add-in specifies a separate image that's more appropriate for the `fr-fr` locale and culture. Users see the image DefaultLogo.png by default, except when the locale of the client computer is `fr-fr`. In this case, users would see the image FrenchLogo.png.</span></span> 


```xml
<!-- Replace "domain" with a real web server name and path. -->
<IconUrl DefaultValue="https://<domain>/DefaultLogo.png"/>
<Override Locale="fr-fr" Value="https://<domain>/FrenchLogo.png"/>
```

<span data-ttu-id="2f827-p116">O exemplo a seguir mostra como localizar um recurso na seção `Resources`. Ele aplica um substituto local para uma imagem que é mais apropriada para a cultura `ja-jp`.</span><span class="sxs-lookup"><span data-stu-id="2f827-p116">The following example shows how to localize a resource in the `Resources` section. It applies a locale override for an image that is more appropriate for the `ja-jp` culture.</span></span>

```xml
<Resources>
      <bt:Images>
        <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
        </bt:Image>
 ...
```


<span data-ttu-id="2f827-p117">Para o elemento [SourceLocation], o suporte a localidades adicionais significa fornecer um arquivo HTML de origem separado para cada um dos locais especificados. Os usuários de cada localidade que você especificar poderão ver uma página da Web personalizada que foi projetada para eles.</span><span class="sxs-lookup"><span data-stu-id="2f827-p117">For the [SourceLocation] element, supporting additional locales means providing a separate source HTML file for each of the specified locales. Users in each locale you specify can see a customized webpage that you design for that them.</span></span>

<span data-ttu-id="2f827-p118">Para suplementos do Outlook, o elemento [SourceLocation] também atribui o fator forma, o que permite que você forneça um arquivo HTML de origem localizado e distinto para cada fator de foram correspondente. Você pode especificar um ou mais elementos filho [Override] em cada configuração aplicável ([DesktopSettings], [TabletSettings] ou [PhoneSettings]). O exemplo a seguir mostra os elementos de configurações para fatores de forma de desktop, tablet e smartphone, cada um com um arquivo HTML para a localidade padrão e outro para a localidade francesa.</span><span class="sxs-lookup"><span data-stu-id="2f827-p118">For Outlook add-ins, the [SourceLocation] element also aligns to the form factor. This allows you to provide a separate, localized source HTML file for each corresponding form factor. You can specify one or more [Override] child elements in each applicable settings element ([DesktopSettings], [TabletSettings], or [PhoneSettings]). The following example shows settings elements for the desktop, tablet, and smartphone form factors, each with one HTML file for the default locale and another for the French locale.</span></span>


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

## <a name="match-datetime-format-with-client-locale"></a><span data-ttu-id="2f827-174">Fazer a correspondência entre o formato de data/hora e a localidade do cliente</span><span class="sxs-lookup"><span data-stu-id="2f827-174">Match date/time format with client locale</span></span>

<span data-ttu-id="2f827-p119">Você pode obter a localidade da interface do usuário do aplicativo host usando a propriedade [displayLanguage]. Em seguida, pode exibir valores de data e hora em um formato consistente com a localidade atual do aplicativo host. Uma maneira de fazer isso é preparar um arquivo de recurso que especifica o formato de exibição de data/hora a ser usado em cada localidade com suporte do seu Suplemento do Office. Na execução, seu suplemento pode usar o arquivo de recurso e fazer a correspondência entre o formato de data/hora apropriado e a localidade obtida na propriedade [displayLanguage]</span><span class="sxs-lookup"><span data-stu-id="2f827-p119">You can get the locale of the user interface of the hosting application by using the [displayLanguage] property. You can then display date and time values in a format consistent with the current locale of the host application. One way to do that is to prepare a resource file that specifies the date/time display format to use for each locale that your Office Add-in supports. At run time, your add-in can use the resource file and match the appropriate date/time format with the locale obtained from the [displayLanguage] property.</span></span>

<span data-ttu-id="2f827-p120">Você pode obter a localidade dos dados do aplicativo host usando a propriedade [contentLanguage]. Com base nesse valor, você pode, então, interpretar ou exibir adequadamente as cadeias de caracteres de data/hora. Por exemplo, a localidade `jp-JP` expressa valores de data/hora como `yyyy/MM/dd`, e a localidade `fr-FR` como `dd/MM/yyyy`.</span><span class="sxs-lookup"><span data-stu-id="2f827-p120">You can get the locale of the data of the hosting application by using the [contentLanguage] property. Based on this value, you can then appropriately interpret or display date/time strings. For example, the `jp-JP` locale expresses data/time values as `yyyy/MM/dd`, and the `fr-FR` locale, `dd/MM/yyyy`.</span></span>


## <a name="use-ajax-for-globalization-and-localization"></a><span data-ttu-id="2f827-182">Usar o Ajax para a globalização e a localização</span><span class="sxs-lookup"><span data-stu-id="2f827-182">Use Ajax for globalization and localization</span></span>


<span data-ttu-id="2f827-183">Se você usar o Visual Studio para criar Suplementos do Office, o .NET Framework e Ajax fornecem maneiras de globalizar e localizar arquivos de script de cliente.</span><span class="sxs-lookup"><span data-stu-id="2f827-183">If you use Visual Studio to create Office Add-ins, the .NET Framework and Ajax provide ways to globalize and localize client script files.</span></span>

<span data-ttu-id="2f827-p121">Você pode globalizar e utilizar as extensões do tipo JavaScript de [Data](https://msdn.microsoft.com/library/caf98d32-2de2-4704-8198-692350343681.aspx) e [Número](https://msdn.microsoft.com/library/c216d3a1-12ae-47d1-bca1-c3666d04572f.aspx) e o objeto [Data](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) do JavaScript no código do JavaScript para um suplemento do Office para exibir valores com base nas configurações de localização do navegador atual. Para saber mais, confira [Passo a passo: como globalizar uma data usando o script de cliente](https://msdn.microsoft.com/library/69b34e6d-d590-4d03-a763-b7ae54b47d74.aspx).</span><span class="sxs-lookup"><span data-stu-id="2f827-p121">You can globalize and use the [Date](https://msdn.microsoft.com/library/caf98d32-2de2-4704-8198-692350343681.aspx) and [Number](https://msdn.microsoft.com/library/c216d3a1-12ae-47d1-bca1-c3666d04572f.aspx) JavaScript type extensions and the JavaScript [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object in the JavaScript code for an Office Add-in to display values based on the locale settings on the current browser. For more information, see [Walkthrough: Globalizing a Date by Using Client Script](https://msdn.microsoft.com/library/69b34e6d-d590-4d03-a763-b7ae54b47d74.aspx).</span></span>

<span data-ttu-id="2f827-p122">Você pode incluir cadeias de caracteres de recurso localizadas diretamente em arquivos de JavaScript autônomos para fornecer arquivos de script de cliente para diferentes locais, que são definidos no navegador ou fornecidos pelo usuário. Crie um arquivo de script separado para cada localidade com suporte. Em cada arquivo de script, inclua um objeto no formato JSON que contenha as cadeias de caracteres de recursos para essa localidade. Os valores localizados serão aplicados quando o script for executado no navegador.</span><span class="sxs-lookup"><span data-stu-id="2f827-p122">You can include localized resource strings directly in standalone JavaScript files to provide client script files for different locales, which are set on the browser or provided by the user. Create a separate script file for each supported locale. In each script file, include an object in JSON format that contains the resource strings for that locale. The localized values are applied when the script runs in the browser.</span></span>


## <a name="example-build-a-localized-office-add-in"></a><span data-ttu-id="2f827-190">Exemplo: Criar um Suplemento do Office localizado</span><span class="sxs-lookup"><span data-stu-id="2f827-190">Example: Build a localized Office Add-in</span></span>

<span data-ttu-id="2f827-191">Esta seção fornece exemplos que mostram como localizar uma descrição do Suplemento do Office, o nome de exibição e interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="2f827-191">This section provides examples that show you how to localize an Office Add-in description, display name, and UI.</span></span> 

> [!NOTE]
> <span data-ttu-id="2f827-192">Para baixar o Visual Studio 2019, confira a [página IDE do Visual Studio](https://visualstudio.microsoft.com/vs/).</span><span class="sxs-lookup"><span data-stu-id="2f827-192">To download Visual Studio 2019, see the [Visual Studio IDE page](https://visualstudio.microsoft.com/vs/).</span></span> <span data-ttu-id="2f827-193">Durante a instalação, você precisará selecionar a carga de trabalho de desenvolvimento do Office/SharePoint.</span><span class="sxs-lookup"><span data-stu-id="2f827-193">During installation you'll need to select the Office/SharePoint development workload.</span></span>

### <a name="configure-office-to-use-additional-languages-for-display-or-editing"></a><span data-ttu-id="2f827-194">Configurar o Office para usar idiomas adicionais para exibição ou edição</span><span class="sxs-lookup"><span data-stu-id="2f827-194">Configure Office to use additional languages for display or editing</span></span>

<span data-ttu-id="2f827-195">Para executar o código de amostra fornecido, configure o Microsoft Office em seu computador para usar idiomas adicionais para que você possa testar seu suplemento, alternando o idioma usado para exibição em menus e em comandos para edição e revisão de texto ou ambos.</span><span class="sxs-lookup"><span data-stu-id="2f827-195">To run the sample code provided, configure Microsoft Office on your computer to use additional languages so that you can test your add-in by switching the language used for display in menus and commands, for editing and proofing, or both.</span></span>

<span data-ttu-id="2f827-196">Você pode usar um Office Language Pack para instalar um idioma adicional.</span><span class="sxs-lookup"><span data-stu-id="2f827-196">You can use an Office Language pack to install an additional language.</span></span> <span data-ttu-id="2f827-197">Para saber mais sobre os Pacotes de Idiomas e onde obtê-los, veja [Language Accessory Pack do Office](https://office.microsoft.com/language-packs/).</span><span class="sxs-lookup"><span data-stu-id="2f827-197">For more information about Language Packs and where to get them, see [Language Accessory Pack for Office](https://office.microsoft.com/language-packs/).</span></span>

<span data-ttu-id="2f827-198">Depois de instalar o Language Accessory Pack, você pode configurar o Office para usar o idioma instalado para exibir na interface do usuário, para edição de conteúdo do documento ou ambos.</span><span class="sxs-lookup"><span data-stu-id="2f827-198">After you install the Language Accessory Pack, you can configure Office to use the installed language for display in the UI, for editing document content, or both.</span></span> <span data-ttu-id="2f827-199">O exemplo neste artigo usa uma instalação do Office que tenha o Pacote de Idiomas de espanhol aplicado.</span><span class="sxs-lookup"><span data-stu-id="2f827-199">The example in this article uses an installation of Office that has the Spanish Language Pack applied.</span></span>

### <a name="create-an-office-add-in-project"></a><span data-ttu-id="2f827-200">Criar um projeto de Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="2f827-200">Create an Office Add-in project</span></span>

<span data-ttu-id="2f827-201">Você precisará criar um projeto de suplemento do Office do Visual Studio 2019.</span><span class="sxs-lookup"><span data-stu-id="2f827-201">You'll need to create a Visual Studio 2019 Office Add-in project.</span></span>

> [!NOTE]
> <span data-ttu-id="2f827-202">Se você ainda não instalou o Visual Studio 2019, consulte a [página do Visual Studio IDE](https://visualstudio.microsoft.com/vs/) para obter instruções de download.</span><span class="sxs-lookup"><span data-stu-id="2f827-202">If you haven't installed Visual Studio 2019, see the [Visual Studio IDE page](https://visualstudio.microsoft.com/vs/) for download instructions.</span></span> <span data-ttu-id="2f827-203">Durante a instalação, você precisará selecionar a carga de trabalho de desenvolvimento do Office/SharePoint.</span><span class="sxs-lookup"><span data-stu-id="2f827-203">During installation you'll need to select the Office/SharePoint development workload.</span></span> <span data-ttu-id="2f827-204">Se você já instalou o Visual Studio 2019, [use o instalador do Visual Studio](/visualstudio/install/modify-visual-studio/) para garantir que a carga de trabalho de desenvolvimento do Office/SharePoint esteja instalada.</span><span class="sxs-lookup"><span data-stu-id="2f827-204">If you have previously installed Visual Studio 2019, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio/) to ensure that the Office/SharePoint development workload is installed.</span></span>

1. <span data-ttu-id="2f827-205">Escolha **criar um novo projeto**.</span><span class="sxs-lookup"><span data-stu-id="2f827-205">Choose **Create a new project**.</span></span>

2. <span data-ttu-id="2f827-206">Usando a caixa de pesquisa, insira o **suplemento**.</span><span class="sxs-lookup"><span data-stu-id="2f827-206">Using the search box, enter **add-in**.</span></span> <span data-ttu-id="2f827-207">Escolha **suplemento da Web do Word**e, em seguida, selecione **Avançar**.</span><span class="sxs-lookup"><span data-stu-id="2f827-207">Choose **Word Web Add-in**, then select **Next**.</span></span>

3. <span data-ttu-id="2f827-208">Nomeie o projeto **WorldReadyAddIn** e selecione **criar**.</span><span class="sxs-lookup"><span data-stu-id="2f827-208">Name your project **WorldReadyAddIn** and select **Create**.</span></span>

4. <span data-ttu-id="2f827-209">O Visual Studio cria uma solução, e os dois projetos dele são exibidos no **Gerenciador de Soluções**.</span><span class="sxs-lookup"><span data-stu-id="2f827-209">Visual Studio creates a solution and its two projects appear in **Solution Explorer**.</span></span> <span data-ttu-id="2f827-210">O arquivo **Home.html** é aberto no Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="2f827-210">The **Home.html** file opens in Visual Studio.</span></span>


### <a name="localize-the-text-used-in-your-add-in"></a><span data-ttu-id="2f827-211">Localizar o texto usado no seu suplemento</span><span class="sxs-lookup"><span data-stu-id="2f827-211">Localize the text used in your add-in</span></span>

<span data-ttu-id="2f827-212">O texto que você deseja localizar para outro idioma aparece em duas áreas:</span><span class="sxs-lookup"><span data-stu-id="2f827-212">The text that you want to localize for another language appears in two areas:</span></span>

-  <span data-ttu-id="2f827-p129">**Nome de exibição e descrição do suplemento**. São controlados por entradas no arquivo de manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="2f827-p129">**Add-in display name and description**. This is controlled by entries in the add-in manifest file.</span></span>

-  <span data-ttu-id="2f827-215">**Interface do usuário do suplemento**.</span><span class="sxs-lookup"><span data-stu-id="2f827-215">**Add-in UI**.</span></span> <span data-ttu-id="2f827-216">Você pode localizar as cadeias de caracteres que aparecem na interface do usuário do seu suplemento usando códigos do JavaScript, por exemplo, usando um arquivo de recurso separado que contenha as cadeias de caracteres localizadas.</span><span class="sxs-lookup"><span data-stu-id="2f827-216">You can localize the strings that appear in your add-in UI by using JavaScript code, for example, by using a separate resource file that contains the localized strings.</span></span>

<span data-ttu-id="2f827-217">Para localizar o nome de exibição e a descrição do suplemento:</span><span class="sxs-lookup"><span data-stu-id="2f827-217">To localize the add-in display name and description:</span></span>

1. <span data-ttu-id="2f827-218">Em **Gerenciador de Soluções**, expanda **WorldReadyAddIn**, **WorldReadyAddInManifest** e, em seguida, selecione **WorldReadyAddIn.xml**.</span><span class="sxs-lookup"><span data-stu-id="2f827-218">In **Solution Explorer**, expand **WorldReadyAddIn**, **WorldReadyAddInManifest**, and then choose  **WorldReadyAddIn.xml**.</span></span>

2. <span data-ttu-id="2f827-219">No arquivo WorldReadyAddInManifest.xml, substitua os elementos [DisplayName] e [Description] pelo seguinte bloco de código:</span><span class="sxs-lookup"><span data-stu-id="2f827-219">In WorldReadyAddInManifest.xml, replace the [DisplayName] and [Description] elements with the following block of code:</span></span>

    > [!NOTE]
    > <span data-ttu-id="2f827-220">Você pode substituir as cadeias de caracteres do idioma espanhol localizado usadas neste exemplo dos elementos [DisplayName] e [Description] pelas cadeias de caracteres localizados para qualquer outro idioma.</span><span class="sxs-lookup"><span data-stu-id="2f827-220">You can replace the Spanish language localized strings used in this example for the [DisplayName] and [Description] elements with the localized strings for any other language.</span></span>

    ```xml
    <DisplayName DefaultValue="World Ready add-in">
      <Override Locale="es-es" Value="Aplicación de uso internacional"/>
    </DisplayName>
    <Description DefaultValue="An add-in for testing localization">
      <Override Locale="es-es" Value="Una aplicación para la prueba de la localización"/>
    </Description>
    ```

3. <span data-ttu-id="2f827-221">Quando você altera o idioma de exibição do Office 2013 do inglês para o espanhol, por exemplo, e executa o suplemento, o nome de exibição do suplemento e a descrição são mostrados com texto localizado.</span><span class="sxs-lookup"><span data-stu-id="2f827-221">When you change the display language for Office 2013 from English to Spanish, for example, and then run the add-in, the add-in display name and description are shown with localized text.</span></span>

<span data-ttu-id="2f827-222">Para definir a interface do usuário do suplemento:</span><span class="sxs-lookup"><span data-stu-id="2f827-222">To lay out the add-in UI:</span></span>

1. <span data-ttu-id="2f827-223">No Visual Studio, no **Gerenciador de Soluções**, selecione **Home.html**.</span><span class="sxs-lookup"><span data-stu-id="2f827-223">In Visual Studio, in **Solution Explorer**, choose **Home.html**.</span></span>

2. <span data-ttu-id="2f827-224">Substitua o conteúdo do elemento `<body>` no Home.html com o HTML a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="2f827-224">Replace the `<body>` element contents in Home.html with the following HTML, and save the file.</span></span>

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

<span data-ttu-id="2f827-225">A figura a seguir mostra o elemento do cabeçalho (h1) e o elemento do parágrafo (p) que exibirá o texto localizado quando concluir as etapas restantes e executar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="2f827-225">The following figure shows the heading (h1) element and the paragraph (p) element that will display localized text when you complete the remaining steps and run the add-in.</span></span>

<span data-ttu-id="2f827-226">*Figura 1. A interface do usuário do suplemento*</span><span class="sxs-lookup"><span data-stu-id="2f827-226">*Figure 1. The add-in UI*</span></span>

![Interface de usuário do aplicativo com as seções realçadas.](../images/office15-app-how-to-localize-fig03.png)

### <a name="add-the-resource-file-that-contains-the-localized-strings"></a><span data-ttu-id="2f827-228">Adicionar o arquivo de recurso que contém as cadeias de caracteres localizadas</span><span class="sxs-lookup"><span data-stu-id="2f827-228">Add the resource file that contains the localized strings</span></span>

<span data-ttu-id="2f827-229">O arquivo de recurso do JavaScript contém as cadeias de caracteres usadas para a interface do usuário do suplemento.</span><span class="sxs-lookup"><span data-stu-id="2f827-229">The JavaScript resource file contains the strings used for the add-in UI.</span></span> <span data-ttu-id="2f827-230">O HTML da interface do usuário do suplemento de amostra tem um elemento `<h1>` que exibe uma saudação e um elemento `<p>` que apresenta o suplemento ao usuário.</span><span class="sxs-lookup"><span data-stu-id="2f827-230">The HTML for the sample add-in UI contains an `<h1>` element that displays a greeting, and a `<p>` element that introduces the add-in to the user.</span></span> 

<span data-ttu-id="2f827-p132">Para habilitar cadeias de caracteres para o cabeçalho e parágrafo, coloque as cadeias de caracteres em um arquivo de recurso separado. O arquivo de recurso cria um objeto do JavaScript que contém um objeto JSON (JavaScript Object Notation) separado para cada conjunto de cadeias de caracteres localizadas. O arquivo de recurso também fornece um método para obter o objeto JSON apropriado de volta para uma determinada localidade.</span><span class="sxs-lookup"><span data-stu-id="2f827-p132">To enable localized strings for the heading and paragraph, you place the strings in a separate resource file. The resource file creates a JavaScript object that contains a separate JavaScript Object Notation (JSON) object for each set of localized strings. The resource file also provides a method for getting back the appropriate JSON object for a given locale.</span></span>

<span data-ttu-id="2f827-234">Para adicionar o arquivo de recurso ao projeto do suplemento:</span><span class="sxs-lookup"><span data-stu-id="2f827-234">To add the resource file to the add-in project:</span></span>

1. <span data-ttu-id="2f827-235">No **Gerenciador de Soluções** no Visual Studio, clique com o botão direito no projeto **WorldReadyAddInWeb** e escolha **Adicionar** > **Novo Item**.</span><span class="sxs-lookup"><span data-stu-id="2f827-235">In **Solution Explorer** in Visual Studio, right-click the **WorldReadyAddInWeb** project and choose **Add** > **New Item**.</span></span> 

2. <span data-ttu-id="2f827-236">Na caixa de diálogo **Adicionar Novo Item**, escolha **Arquivo JavaScript**.</span><span class="sxs-lookup"><span data-stu-id="2f827-236">In the **Add New Item** dialog box, choose **JavaScript File**.</span></span>

3. <span data-ttu-id="2f827-237">Insira **UIStrings.js** como nome do arquivo e escolha **Adicionar**.</span><span class="sxs-lookup"><span data-stu-id="2f827-237">Enter **UIStrings.js** as the file name and choose **Add**.</span></span>

4. <span data-ttu-id="2f827-238">Adicione o código a seguir ao arquivo UIStrings.js e salve-o.</span><span class="sxs-lookup"><span data-stu-id="2f827-238">Add the following code to the UIStrings.js file, and save the file.</span></span>

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

<span data-ttu-id="2f827-239">O arquivo de recurso UIStrings.js cria o objeto, **UIStrings**, que contém as cadeias de caracteres localizadas para a interface do usuário do suplemento.</span><span class="sxs-lookup"><span data-stu-id="2f827-239">The UIStrings.js resource file creates an object, **UIStrings**, which contains the localized strings for your add-in UI.</span></span>

### <a name="localize-the-text-used-for-the-add-in-ui"></a><span data-ttu-id="2f827-240">Localizar o texto usado na interface do usuário do suplemento</span><span class="sxs-lookup"><span data-stu-id="2f827-240">Localize the text used for the add-in UI</span></span>

<span data-ttu-id="2f827-p133">Para usar o arquivo de recurso no seu suplemento, você precisará adicionar a ele uma marca de script em Home.html. Quando Home.html for carregado, o UIStrings.js será executado e o objeto **UIStrings** que você utiliza para obter a cadeia de caracteres ficará disponível para seu código. Adicione o seguinte HTML à marca de cabeçalho do Home.html para tornar **UIStrings** disponível para seu código.</span><span class="sxs-lookup"><span data-stu-id="2f827-p133">To use the resource file in your add-in, you'll need to add a script tag for it on Home.html. When Home.html is loaded, UIStrings.js executes and the **UIStrings** object that you use to get the strings is available to your code. Add the following HTML in the head tag for Home.html to make **UIStrings** available to your code.</span></span>

```html
<!-- Resource file for localized strings:                                                          -->
<script src="../UIStrings.js" type="text/javascript"></script>
```

<span data-ttu-id="2f827-244">Agora você pode usar o objeto **UIStrings** para definir as cadeias de caracteres da interface do usuário do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="2f827-244">Now you can use the **UIStrings** object to set the strings for the UI of your add-in.</span></span>

<span data-ttu-id="2f827-p134">Se você quiser alterar a localização do seu suplemento com base no idioma usado para exibição nos menus e comandos no aplicativo host, use a propriedade **Office.context.displayLanguage** para obter a localidade desse idioma. Por exemplo, se o idioma do aplicativo host utilizar espanhol para exibir menus e comandos, a propriedade **Office.context.displayLanguage** retornará o código es-ES.</span><span class="sxs-lookup"><span data-stu-id="2f827-p134">If you want to change the localization for your add-in based on what language is used for display in menus and commands in the host application, you use the **Office.context.displayLanguage** property to get the locale for that language. For example, if the host application language uses Spanish for display in menus and commands, the **Office.context.displayLanguage** property will return the language code es-ES.</span></span>

<span data-ttu-id="2f827-p135">Se você quiser alterar a localização do seu suplemento com base no idioma que está sendo usado para editar o conteúdo do documento, use a propriedade **Office.context.contentLanguage** para obter a localidade do idioma. Por exemplo, se o idioma do aplicativo host utilizar espanhol para editar o conteúdo do documento, a propriedade **Office.context.contentLanguage** retornará o código es-ES.</span><span class="sxs-lookup"><span data-stu-id="2f827-p135">If you want to change the localization for your add-in based on what language is being used for editing document content, you use the  **Office.context.contentLanguage** property to get the locale for that language. For example, if the host application language uses Spanish for editing document content, the **Office.context.contentLanguage** property will return the language code es-ES.</span></span>

<span data-ttu-id="2f827-249">Depois que você souber o idioma que o aplicativo host está utilizando, é possível usar **UIStrings** para obter o conjunto de cadeias de caracteres localizadas correspondentes ao idioma do aplicativo host.</span><span class="sxs-lookup"><span data-stu-id="2f827-249">After you know the language the host application is using, you can use **UIStrings** to get the set of localized strings that matches the host application language.</span></span>

<span data-ttu-id="2f827-p136">Substitua o código no arquivo Home.js pelo código a seguir. O código mostra como você pode alterar as cadeias de caracteres usadas nos elementos da interface do usuário no Home.html com base no idioma de exibição do aplicativo host ou no idioma de edição do aplicativo host.</span><span class="sxs-lookup"><span data-stu-id="2f827-p136">Replace the code in the Home.js file with the following code. The code shows how you can change the strings used in the UI elements on Home.html based on either the display language of the host application or the editing language of the host application.</span></span>

> [!NOTE]
> <span data-ttu-id="2f827-252">Para alternar entre a alteração da localização do suplemento com base no idioma usado para edição, remova o comentário da linha de código `var myLanguage = Office.context.contentLanguage;` e inclua o comentário na linha de código `var myLanguage = Office.context.displayLanguage;`</span><span class="sxs-lookup"><span data-stu-id="2f827-252">To switch between changing the localization of the add-in based on the language used for editing, uncomment the line of code  `var myLanguage = Office.context.contentLanguage;` and comment out the line of code `var myLanguage = Office.context.displayLanguage;`</span></span>

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

            // Get the language setting for UI display in the host application.
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

### <a name="test-your-localized-add-in"></a><span data-ttu-id="2f827-253">Testar seu suplemento localizado</span><span class="sxs-lookup"><span data-stu-id="2f827-253">Test your localized add-in</span></span>

<span data-ttu-id="2f827-254">Para testar seu suplemento localizado, altere o idioma usado para exibir ou editar no aplicativo host e execute o seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="2f827-254">To test your localized add-in, change the language used for display or editing in the host application and then run your add-in.</span></span>

<span data-ttu-id="2f827-255">Para alterar o idioma usado para exibir ou editar no seu suplemento:</span><span class="sxs-lookup"><span data-stu-id="2f827-255">To change the language used for display or editing in your add-in:</span></span>

1. <span data-ttu-id="2f827-256">No Word, escolha **Arquivo** > **Opções** > **Idioma**.</span><span class="sxs-lookup"><span data-stu-id="2f827-256">In Word, choose **File** > **Options** > **Language**.</span></span> <span data-ttu-id="2f827-257">A figura a seguir mostra a caixa de diálogo **Opções do Word** aberta na guia Idioma.</span><span class="sxs-lookup"><span data-stu-id="2f827-257">The following figure shows the **Word Options** dialog box opened to the Language tab.</span></span>

    <span data-ttu-id="2f827-258">*Figura 2. Opções de idioma na caixa de diálogo Opções do Word*</span><span class="sxs-lookup"><span data-stu-id="2f827-258">*Figure 2. Language options in the Word Options dialog box*</span></span>

    ![Caixa de diálogo Opções do Word](../images/office15-app-how-to-localize-fig04.png)

2. <span data-ttu-id="2f827-260">Em **Escolher Idioma de Exibição**, selecione o idioma desejado para exibição, por exemplo, espanhol, e selecione a seta para cima para mover o idioma espanhol para a primeira posição na lista.</span><span class="sxs-lookup"><span data-stu-id="2f827-260">Under **Choose Display Language**, select the language that you want for display, for example Spanish, and then choose the up arrow to move the Spanish language to the first position in the list.</span></span> <span data-ttu-id="2f827-261">Ou, para alterar o idioma usado para edição, em **Escolher Idiomas de Edição**, escolha o idioma que você deseja usar para edição, por exemplo, espanhol, e selecione **Definir como Padrão**.</span><span class="sxs-lookup"><span data-stu-id="2f827-261">Alternatively, to change the language used for editing, under  **Choose Editing Languages**, choose the language you want to use for editing, for example, Spanish, and then choose **Set as Default**.</span></span>

3. <span data-ttu-id="2f827-262">Escolha **OK** para confirmar sua seleção e feche o Word.</span><span class="sxs-lookup"><span data-stu-id="2f827-262">Choose **OK** to confirm your selection, and then close Word.</span></span>

4. <span data-ttu-id="2f827-263">Pressione **F5** no Visual Studio para executar o suplemento de amostra ou escolha **Depurar** > **Iniciar Depuração** na barra de menus.</span><span class="sxs-lookup"><span data-stu-id="2f827-263">Press **F5** in Visual Studio to run the sample add-in, or choose **Debug** > **Start Debugging** from the menu bar.</span></span>

5. <span data-ttu-id="2f827-264">No Word, escolha **Página Inicial** > **Mostrar Painel de Tarefas**.</span><span class="sxs-lookup"><span data-stu-id="2f827-264">In Word, choose **Home** > **Show Taskpane**.</span></span>

<span data-ttu-id="2f827-265">Após estar em execução, as cadeias de caracteres na interface do usuário do suplemento são alteradas para corresponder ao idioma usado pelo aplicativo host, como mostrado na figura a seguir.</span><span class="sxs-lookup"><span data-stu-id="2f827-265">Once running, the strings in the add-in UI change to match the language used by the host application, as shown in the following figure.</span></span>


<span data-ttu-id="2f827-266">*Figura 3. Interface do usuário do suplemento com o texto localizado*</span><span class="sxs-lookup"><span data-stu-id="2f827-266">*Figure 3. Add-in UI with localized text*</span></span>

![Aplicativo com texto localizado da interface do usuário.](../images/office15-app-how-to-localize-fig05.png)

## <a name="see-also"></a><span data-ttu-id="2f827-268">Confira também</span><span class="sxs-lookup"><span data-stu-id="2f827-268">See also</span></span>

- [<span data-ttu-id="2f827-269">Diretrizes de design para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="2f827-269">Design guidelines for Office Add-ins</span></span>](../design/add-in-design.md)
- <span data-ttu-id="2f827-270">[Identificadores de idioma e valores da ID de OptionState no Office 2013](/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15))</span><span class="sxs-lookup"><span data-stu-id="2f827-270">[Language identifiers and OptionState Id values in Office 2013](/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15))</span></span>

[DefaultLocale]:        /office/dev/add-ins/reference/manifest/defaultlocale
[Descrição]:          /office/dev/add-ins/reference/manifest/description
[Description]:          /office/dev/add-ins/reference/manifest/description
[DisplayName]:          /office/dev/add-ins/reference/manifest/displayname
[IconUrl]:              /office/dev/add-ins/reference/manifest/iconurl
[HighResolutionIconUrl]:/office/dev/add-ins/reference/manifest/highresolutioniconurl
[Resources]:            /office/dev/add-ins/reference/manifest/resources
[SourceLocation]:       /office/dev/add-ins/reference/manifest/sourcelocation
[Override]:             /office/dev/add-ins/reference/manifest/override
[DesktopSettings]:      /office/dev/add-ins/reference/manifest/desktopsettings
[TabletSettings]:       /office/dev/add-ins/reference/manifest/tabletsettings
[PhoneSettings]:        /office/dev/add-ins/reference/manifest/phonesettings
[displayLanguage]:  /javascript/api/office/office.context#displaylanguage 
[contentLanguage]:  /javascript/api/office/office.context#contentlanguage 
[RFC 3066]: https://www.rfc-editor.org/info/rfc3066
