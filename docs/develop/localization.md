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
# <a name="localization-for-office-add-ins"></a><span data-ttu-id="9f5f2-102">Localiza??o para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="9f5f2-102">Localization for Office Add-ins</span></span>

<span data-ttu-id="9f5f2-p101">Voc? pode implementar qualquer esquema de localiza??o que seja apropriado para o seu Suplemento do Office. A API JavaScript e o esquema do manifesto da plataforma de Suplementos do Office oferecem algumas op??es. Voc? pode usar a API JavaScript para Office para determinar uma localidade e exibir cadeias de caracteres com base na localidade do aplicativo host ou para interpretar ou exibir dados com base na localidade dos dados. Voc? pode usar o manifesto para especificar informa??es descritivas e o local do arquivo do suplemento espec?fico da localidade. Como alternativa, voc? pode usar o script do Microsoft Ajax para dar suporte ? globaliza??o e localiza??o.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p101">You can implement any localization scheme that's appropriate for your Office Add-in. The JavaScript API and manifest schema of the Office Add-ins platform provide some choices. You can use the JavaScript API for Office to determine a locale and display strings based on the locale of the host application, or to interpret or display data based on the locale of the data. You can use the manifest to specify locale-specific add-in file location and descriptive information. Alternatively, you can use Microsoft Ajax script to support globalization and localization.</span></span>

## <a name="use-the-javascript-api-to-determine-locale-specific-strings"></a><span data-ttu-id="9f5f2-108">Usar a API JavaScript para determinar cadeias de caracteres espec?ficas da localidade</span><span class="sxs-lookup"><span data-stu-id="9f5f2-108">Use the JavaScript API to determine locale-specific strings</span></span>

<span data-ttu-id="9f5f2-109">A API JavaScript para Office fornece duas propriedades que oferecem suporte ? exibi??o ou interpreta??o de valores consistentes com a localidade do aplicativo host e dos dados:</span><span class="sxs-lookup"><span data-stu-id="9f5f2-109">The JavaScript API for Office provides two properties that support displaying or interpreting values consistent with the locale of the host application and data:</span></span>

- <span data-ttu-id="9f5f2-p102">[Context.displayLanguage][displayLanguage] especifica a localidade (ou idioma) da interface do usu?rio do aplicativo host. O exemplo a seguir verifica se o aplicativo host usa a localidade en-US ou fr-FR e exibe uma sauda??o espec?fica para a localidade.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p102">[Context.displayLanguage][displayLanguage] specifies the locale (or language) of the user interface of the host application. The following example verifies if the host application uses the en-US or fr-Fr locale, and displays a locale-specific greeting.</span></span>
    
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

- <span data-ttu-id="9f5f2-p103">[Context.contentLanguage][contentLanguage] especifica a localidade (ou o idioma) dos dados. Estendendo o ?ltimo exemplo de c?digo, em vez de verificar a propriedade [displayLanguage], atribua `myLanguage` ? propriedade [contentLanguage] e use o restante do mesmo c?digo para exibir uma sauda??o com base na localidade dos dados:</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p103">[Context.contentLanguage][contentLanguage] specifies the locale (or language) of the data. Extending the last code sample, instead of checking the [displayLanguage] property, assign `myLanguage` to the [contentLanguage] property, and use the rest of the same code to display a greeting based on the locale of the data:</span></span>
    
    ```js
    var myLanguage = Office.context.contentLanguage;
    ```

## <a name="control-localization-from-the-manifest"></a><span data-ttu-id="9f5f2-114">Controlar a localiza??o do manifesto</span><span class="sxs-lookup"><span data-stu-id="9f5f2-114">Control localization from the manifest</span></span>


<span data-ttu-id="9f5f2-p104">Cada Suplemento do Office especifica um elemento [DefaultLocale] e uma localidade em seu manifesto. Por padr?o, a plataforma do Suplemento do Office e os aplicativos host do Office aplicam os valores dos elementos [Description], [DisplayName], [IconUrl], [HighResolutionIconUrl] e [SourceLocation] a todas as localidades. Como op??o, voc? pode dar suporte a valores espec?ficos para localidades espec?ficas, especificando um elemento-filho [Override]para cada localidade adicional, para qualquer um desses cinco elementos. O valor do elemento [DefaultLocale] e do atributo `Locale` do elemento [Override] ? especificado de acordo com o [RFC 3066], "Marcas para a Identifica??o dos Idiomas". A Tabela 1 descreve o suporte de localiza??o para esses elementos.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p104">Every Office Add-in specifies a [DefaultLocale] element and a locale in its manifest. By default, the Office Add-in platform and Office host applications apply the values of the [Description], [DisplayName], [IconUrl], [HighResolutionIconUrl], and [SourceLocation] elements to all locales. You can optionally support specific values for specific locales, by specifying an [Override] child element for each additional locale, for any of these five elements. The value for the [DefaultLocale] element and for the `Locale` attribute of the [Override] element is specified according to [RFC 3066], "Tags for the Identification of Languages." Table 1 describes the localizing support for these elements.</span></span>

<span data-ttu-id="9f5f2-120">**Tabela 1. Suporte de localiza??o**</span><span class="sxs-lookup"><span data-stu-id="9f5f2-120">**Table 1. Localization support**</span></span>


|<span data-ttu-id="9f5f2-121">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="9f5f2-121">**Element**</span></span>|<span data-ttu-id="9f5f2-122">**Suporte de localiza??o**</span><span class="sxs-lookup"><span data-stu-id="9f5f2-122">**Localization support**</span></span>|
|:-----|:-----|
|<span data-ttu-id="9f5f2-123">[Descri??o]</span><span class="sxs-lookup"><span data-stu-id="9f5f2-123">[Description]</span></span>   |<span data-ttu-id="9f5f2-124">Os usu?rios de cada localidade especificada podem ver uma descri??o localizada do suplemento no AppSource (ou no cat?logo privado).</span><span class="sxs-lookup"><span data-stu-id="9f5f2-124">Users in each locale you specify can see a localized description for the add-in in AppSource (or private catalog).</span></span><br/><span data-ttu-id="9f5f2-125">Para os suplementos do Outlook, os usu?rios podem ver a descri??o no Centro de Administra??o do Exchange (EAC) ap?s a instala??o.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-125">For Outlook add-ins, users can see the description in the Exchange Admin Center (EAC) after installation.</span></span>|
|<span data-ttu-id="9f5f2-126">[DisplayName]</span><span class="sxs-lookup"><span data-stu-id="9f5f2-126">[DisplayName]</span></span>   |<span data-ttu-id="9f5f2-127">Os usu?rios de cada localidade especificada podem ver uma descri??o localizada do suplemento no AppSource (ou no cat?logo privado).</span><span class="sxs-lookup"><span data-stu-id="9f5f2-127">Users in each locale you specify can see a localized description for the add-in in AppSource (or private catalog).</span></span><br/><span data-ttu-id="9f5f2-128">Para os suplementos do Outlook, os usu?rios podem ver o nome de exibi??o como um r?tulo para o bot?o de suplemento do Outlook e no EAC ap?s a instala??o.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-128">For Outlook add-ins, users can see the display name as a label for the Outlook add-in button and in the EAC after installation.</span></span><br/><span data-ttu-id="9f5f2-129">Para os suplementos do painel de tarefas e do conte?do, os usu?rios podem ver o nome de exibi??o na faixa de op??es ap?s a instala??o do suplemento.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-129">For content and task pane add-ins, users can see the display name in the ribbon after installing the add-in.</span></span>|
|<span data-ttu-id="9f5f2-130">[IconUrl]</span><span class="sxs-lookup"><span data-stu-id="9f5f2-130">[IconUrl]</span></span>        |<span data-ttu-id="9f5f2-p105">A imagem do ?cone ? opcional. Voc? pode usar a mesma t?cnica de substitui??o para especificar uma determinada imagem para uma cultura espec?fica. Se voc? usar e localizar um ?cone, os usu?rios em cada localidade que voc? especificar poder?o ver uma imagem de ?cone localizada para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p105">The icon image is optional. You can use the same override technique to specify a certain image for a specific culture. If you use and localize an icon, users in each locale you specify can see a localized icon image for the add-in.</span></span><br/><span data-ttu-id="9f5f2-134">Para suplementos do Outlook, os usu?rios podem ver o ?cone no EAC depois de instalar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-134">For Outlook add-ins, users can see the icon in the EAC after installing the add-in.</span></span><br/><span data-ttu-id="9f5f2-135">Para os suplementos do painel de tarefas e do conte?do, os usu?rios podem ver o ?cone na faixa de op??es ap?s a instala??o do suplemento.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-135">For content and task pane add-ins, users can see the icon in the ribbon after installing the add-in.</span></span>|
|<span data-ttu-id="9f5f2-136">[HighResolutionIconUrl] **Importante:** este elemento s? fica dispon?vel ao usar a vers?o 1.1 do manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-136">[HighResolutionIconUrl] **Important:** This element is available only when using add-in manifest version 1.1.</span></span>|<span data-ttu-id="9f5f2-p106">A imagem do ?cone de alta resolu??o ? opcional, mas se ela for especificada, dever? ocorrer ap?s o elemento [IconUrl]. Quando [HighResolutionIconUrl] for especificado e o suplemento estiver instalado em um dispositivo que ofere?a suporte ? resolu??o dpi alto, o valor [HighResolutionIconUrl] ? usado em vez do valor para [IconUrl].</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p106">The high resolution icon image is optional but if it is specified, it must occur after the  [IconUrl] element. When [HighResolutionIconUrl] is specified, and the add-in is installed on a device that supports high dpi resolution, the [HighResolutionIconUrl] value is used instead of the value for [IconUrl].</span></span><br/><span data-ttu-id="9f5f2-p107">Voc? pode usar a mesma t?cnica de substitui??o para especificar uma determinada imagem para uma cultura espec?fica. Se voc? usar e localizar um ?cone, os usu?rios em cada localidade que voc? especificar podem ver uma imagem de ?cone localizada para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p107">You can use the same override technique to specify a certain image for a specific culture. If you use and localize an icon, users in each locale you specify can see a localized icon image for the add-in.</span></span><br/><span data-ttu-id="9f5f2-141">Para suplementos do Outlook, os usu?rios podem ver o ?cone no EAC depois de instalar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-141">For Outlook add-ins, users can see the icon in the EAC after installing the add-in.</span></span><br/><span data-ttu-id="9f5f2-142">Para os suplementos do painel de tarefas e do conte?do, os usu?rios podem ver o ?cone na faixa de op??es ap?s a instala??o do suplemento.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-142">For content and task pane add-ins, users can see the icon in the ribbon after installing the add-in.</span></span>|
|<span data-ttu-id="9f5f2-143">[Recursos] **Importante:** este elemento s? fica dispon?vel ao usar a vers?o 1.1 do manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-143">[Resources] **Important:** This element is available only when using add-in manifest version 1.1.</span></span>   |<span data-ttu-id="9f5f2-144">Os usu?rios em cada localidade especificada podem ver recursos de cadeias de caracteres e de ?cones que voc? projetou especificamente para o suplemento dessa localidade.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-144">Users in each locale you specify can see string and icon resources that you specifically create for the add-in for that locale.</span></span> |
|<span data-ttu-id="9f5f2-145">[SourceLocation]</span><span class="sxs-lookup"><span data-stu-id="9f5f2-145">[SourceLocation]</span></span>   |<span data-ttu-id="9f5f2-146">Os usu?rios em cada localidade especificada podem ver uma p?gina da Web que voc? projetou especificamente para o suplemento dessa localidade.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-146">Users in each locale you specify can see a webpage that you specifically design for the add-in for that locale.</span></span> |


> <span data-ttu-id="9f5f2-p108">**OBSERVA??O:** voc? s? pode localizar o nome de exibi??o e a descri??o das localidades que oferecem suporte ao Office. Veja [Identificadores de idioma e valores de OptionState Id no Office 2013](http://technet.microsoft.com/en-us/library/cc179219.aspx) para obter uma lista de idiomas e localidades para a vers?o atual do Office.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p108">**NOTE** You can localize the description and display name for only the locales that Office supports. See [Language identifiers and OptionState Id values in Office 2013](http://technet.microsoft.com/en-us/library/cc179219.aspx) for a list of languages and locales for the current release of Office.</span></span>


### <a name="examples"></a><span data-ttu-id="9f5f2-149">Exemplos</span><span class="sxs-lookup"><span data-stu-id="9f5f2-149">Examples</span></span>

<span data-ttu-id="9f5f2-p109">Por exemplo, um Suplemento do Office pode especificar o [DefaultLocale] como `en-us`. Para o elemento [DisplayName], o suplemento pode especificar um elemento filho [Override] para a localidade `fr-fr`, como mostrado abaixo.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p109">For example, an Office Add-in can specify the [DefaultLocale] as `en-us`. For the [DisplayName] element, the add-in can specify an [Override] child element for the locale `fr-fr`, as shown below.</span></span> 


```xml
<DefaultLocale>en-us</DefaultLocale>
...
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

> <span data-ttu-id="9f5f2-p110">**OBSERVA??O:** se for preciso localizar para mais de uma ?rea dentro de uma fam?lia de idiomas, como `de-de` e `de-at`, recomendamos que voc? use elementos `Override` separados para cada ?rea. Usar apenas o nome do idioma sozinho, nesse caso, `de`, n?o tem suporte em todas as combina??es de plataformas e aplicativos de host do Office.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p110">**NOTE** If you need to localize for more than one area within a language family, such as `de-de` and `de-at`, we recommend that you use separate `Override` elements for each area. Using just the language name alone, in this case, `de`, is not supported across all combinations of Office host applications and platforms.</span></span>

<span data-ttu-id="9f5f2-p111">Isso significa que o suplemento pressup?e a localidade `en-us` como padr?o. Os usu?rios veem o nome de exibi??o em ingl?s "Video player" para todas as localidades, a menos que a localidade do computador cliente fosse `fr-fr`, nesse caso os usu?rios veria o nome de exibi??o em franc?s "Lecteur vid?o".</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p111">This means that the add-in assumes the  `en-us` locale by default. Users see the English display name of "Video player" for all locales unless the client computer's locale is `fr-fr`, in which case users would see the French display name "Lecteur vid?o".</span></span>

> <span data-ttu-id="9f5f2-p112">**Observa??o:** voc? s? pode especificar uma ?nica substitui??o por idioma, inclusive para a localidade padr?o. Por exemplo, se sua localidade padr?o ? `en-us`, n?o ? poss?vel especificar tamb?m uma substitui??o para `en-us`.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p112">**NOTE** You may only specify a single override per language, including for the default locale. For example, if your default locale is `en-us` you cannot not specify an  override for `en-us` as well.</span></span> 

<span data-ttu-id="9f5f2-p113">O exemplo a seguir se aplica a uma substitui??o de localidade para o elemento [Description]. Primeiro especifica a localidade padr?o `en-us` e uma descri??o em ingl?s e, em seguida, especifica uma pol?tica de [Override] com uma descri??o francesa para a localidade `fr-fr`:</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p113">The following example applies a locale override for the [Description] element. It first specifies a default locale of `en-us` and an English description, and then specifies an [Override] statement with a French description for the `fr-fr` locale:</span></span>

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

<span data-ttu-id="9f5f2-p114">Isso significa que o suplemento pressup?e a localidade `en-us` como padr?o. Os usu?rios veriam a descri??o em ingl?s no atributo `DefaultValue` para todas as localidades, a menos que a localidade do computador cliente fosse `fr-fr`, nesse caso, eles veriam a descri??o em franc?s.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p114">This means that the add-in assumes the `en-us` locale by default. Users would see the English description in the `DefaultValue` attribute for all locales unless the client computer's locale is `fr-fr`, in which case they would see the French description.</span></span>

<span data-ttu-id="9f5f2-p115">No exemplo a seguir, o suplemento especifica uma imagem separada mais apropriada para a localidade e a cultura `fr-fr`. Os usu?rios ver?o a imagem DefaultLogo.png por padr?o, exceto quando a localidade do computador cliente for `fr-fr`. Nesse caso, os usu?rios veriam a imagem FrenchLogo.png.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p115">In the following example, the add-in specifies a separate image that's more appropriate for the `fr-fr` locale and culture. Users see the image DefaultLogo.png by default, except when the locale of the client computer is `fr-fr`. In this case, users would see the image FrenchLogo.png.</span></span> 


```xml
<!-- Replace "domain" with a real web server name and path. -->
<IconUrl DefaultValue="https://<domain>/DefaultLogo.png"/>
<Override Locale="fr-fr" Value="https://<domain>/FrenchLogo.png"/>
```

<span data-ttu-id="9f5f2-p116">O exemplo a seguir mostra como localizar um recurso na se??o `Resources`. Ele aplica um substituto local para uma imagem que ? mais apropriada para a cultura `ja-jp`.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p116">The following example shows how to localize a resource in the `Resources` section. It applies a locale override for an image that is more appropriate for the `ja-jp` culture.</span></span>

```xml
<Resources>
      <bt:Images>
        <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
        </bt:Image>
 ...
```


<span data-ttu-id="9f5f2-p117">Para o elemento [SourceLocation], o suporte a localidades adicionais significa fornecer um arquivo HTML de origem separado para cada um dos locais especificados. Os usu?rios de cada localidade que voc? especificar poder?o ver uma p?gina da Web personalizada que foi projetada para eles.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p117">For the [SourceLocation] element, supporting additional locales means providing a separate source HTML file for each of the specified locales. Users in each locale you specify can see a customized webpage that you design for that them.</span></span>

<span data-ttu-id="9f5f2-p118">Para suplementos do Outlook, o elemento [SourceLocation] tamb?m atribui o fator forma, o que permite que voc? forne?a um arquivo HTML de origem localizado e distinto para cada fator de foram correspondente. Voc? pode especificar um ou mais elementos filho [Override] em cada configura??o aplic?vel ([DesktopSettings], [TabletSettings] ou [PhoneSettings]). O exemplo a seguir mostra os elementos de configura??es para fatores de forma de desktop, tablet e smartphone, cada um com um arquivo HTML para a localidade padr?o e outro para a localidade francesa.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p118">For Outlook add-ins, the [SourceLocation] element also aligns to the form factor. This allows you to provide a separate, localized source HTML file for each corresponding form factor. You can specify one or more [Override] child elements in each applicable settings element ([DesktopSettings], [TabletSettings], or [PhoneSettings]). The following example shows settings elements for the desktop, tablet, and smartphone form factors, each with one HTML file for the default locale and another for the French locale.</span></span>


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

## <a name="match-datetime-format-with-client-locale"></a><span data-ttu-id="9f5f2-173">Fazer a correspond?ncia entre o formato de data/hora e a localidade do cliente</span><span class="sxs-lookup"><span data-stu-id="9f5f2-173">Match date/time format with client locale</span></span>

<span data-ttu-id="9f5f2-p119">Voc? pode obter a localidade da interface do usu?rio do aplicativo host usando a propriedade [displayLanguage]. Em seguida, pode exibir valores de data e hora em um formato consistente com a localidade atual do aplicativo host. Uma maneira de fazer isso ? preparar um arquivo de recurso que especifica o formato de exibi??o de data/hora a ser usado em cada localidade com suporte do seu Suplemento do Office. Na execu??o, seu suplemento pode usar o arquivo de recurso e fazer a correspond?ncia entre o formato de data/hora apropriado e a localidade obtida na propriedade [displayLanguage]</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p119">You can get the locale of the user interface of the hosting application by using the [displayLanguage] property. You can then display date and time values in a format consistent with the current locale of the host application. One way to do that is to prepare a resource file that specifies the date/time display format to use for each locale that your Office Add-in supports. At run time, your add-in can use the resource file and match the appropriate date/time format with the locale obtained from the [displayLanguage] property.</span></span>

<span data-ttu-id="9f5f2-p120">Voc? pode obter a localidade dos dados do aplicativo host usando a propriedade [contentLanguage]. Com base nesse valor, voc? pode, ent?o, interpretar ou exibir adequadamente as cadeias de caracteres de data/hora. Por exemplo, a localidade `jp-JP` expressa valores de data/hora como `yyyy/MM/dd`, e a localidade `fr-FR` como `dd/MM/yyyy`.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p120">You can get the locale of the data of the hosting application by using the [contentLanguage] property. Based on this value, you can then appropriately interpret or display date/time strings. For example, the `jp-JP` locale expresses data/time values as `yyyy/MM/dd`, and the `fr-FR` locale, `dd/MM/yyyy`.</span></span>


## <a name="use-ajax-for-globalization-and-localization"></a><span data-ttu-id="9f5f2-181">Usar o Ajax para a globaliza??o e a localiza??o</span><span class="sxs-lookup"><span data-stu-id="9f5f2-181">Use Ajax for globalization and localization</span></span>


<span data-ttu-id="9f5f2-182">Se voc? usar o Visual Studio para criar Suplementos do Office, o .NET Framework e Ajax fornecem maneiras de globalizar e localizar arquivos de script de cliente.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-182">If you use Visual Studio to create Office Add-ins, the .NET Framework and Ajax provide ways to globalize and localize client script files.</span></span>

<span data-ttu-id="9f5f2-p121">Voc? pode globalizar e utilizar as extens?es do tipo JavaScript de [Data](http://msdn.microsoft.com/library/caf98d32-2de2-4704-8198-692350343681.aspx) e [N?mero](http://msdn.microsoft.com/library/c216d3a1-12ae-47d1-bca1-c3666d04572f.aspx) e o objeto [Data](http://msdn.microsoft.com/library/ce2202bb-7ec9-4f5a-bf48-3a04feff283e.aspx) do JavaScript no c?digo do JavaScript para um suplemento do Office para exibir valores com base nas configura??es de localiza??o do navegador atual. Para saber mais, confira [Passo a passo: como globalizar uma data usando o script de cliente](http://msdn.microsoft.com/library/69b34e6d-d590-4d03-a763-b7ae54b47d74.aspx).</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p121">You can globalize and use the [Date](http://msdn.microsoft.com/library/caf98d32-2de2-4704-8198-692350343681.aspx) and [Number](http://msdn.microsoft.com/library/c216d3a1-12ae-47d1-bca1-c3666d04572f.aspx) JavaScript type extensions and the JavaScript [Date](http://msdn.microsoft.com/library/ce2202bb-7ec9-4f5a-bf48-3a04feff283e.aspx) object in the JavaScript code for an Office Add-in to display values based on the locale settings on the current browser. For more information, see [Walkthrough: Globalizing a Date by Using Client Script](http://msdn.microsoft.com/library/69b34e6d-d590-4d03-a763-b7ae54b47d74.aspx).</span></span>

<span data-ttu-id="9f5f2-p122">Voc? pode incluir cadeias de caracteres de recurso localizadas diretamente em arquivos de JavaScript aut?nomos para fornecer arquivos de script de cliente para diferentes locais, que s?o definidos no navegador ou fornecidos pelo usu?rio. Crie um arquivo de script separado para cada localidade com suporte. Em cada arquivo de script, inclua um objeto no formato JSON que contenha as cadeias de caracteres de recursos para essa localidade. Os valores localizados ser?o aplicados quando o script for executado no navegador.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p122">You can include localized resource strings directly in standalone JavaScript files to provide client script files for different locales, which are set on the browser or provided by the user. Create a separate script file for each supported locale. In each script file, include an object in JSON format that contains the resource strings for that locale. The localized values are applied when the script runs in the browser.</span></span> 


## <a name="example-build-a-localized-office-add-in"></a><span data-ttu-id="9f5f2-189">Exemplo: Criar um Suplemento do Office localizado</span><span class="sxs-lookup"><span data-stu-id="9f5f2-189">Example: Build a localized Office Add-in</span></span>

<span data-ttu-id="9f5f2-190">Esta se??o fornece exemplos que mostram como localizar uma descri??o do Suplemento do Office, o nome de exibi??o e interface do usu?rio.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-190">This section provides examples that show you how to localize an Office Add-in description, display name, and UI.</span></span>

<span data-ttu-id="9f5f2-191">Para executar o c?digo de amostra fornecido, configure o Microsoft Office 2013 em seu computador para usar idiomas adicionais para que voc? possa testar seu suplemento, alternando o idioma usado para exibi??o em menus e em comandos para edi??o e revis?o de texto ou ambos.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-191">To run the sample code provided, configure Microsoft Office 2013 on your computer to use additional languages so that you can test your add-in by switching the language used for display in menus and commands, for editing and proofing, or both.</span></span>

<span data-ttu-id="9f5f2-192">Al?m disso, voc? precisar? criar um projeto de Suplemento do Office do Visual Studio 2015.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-192">Also, you'll need to create a Visual Studio 2015 Office Add-in project.</span></span>

> <span data-ttu-id="9f5f2-p123">**Observa??o:** para baixar o Visual Studio 2015, confira a [P?gina do Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs). Esta p?gina tamb?m tem um link para o Office Developer Tools.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p123">**NOTE** To download Visual Studio 2015, see the [Office Developer Tools page](https://www.visualstudio.com/features/office-tools-vs). This page also has a link for the Office Developer Tools.</span></span>

### <a name="configure-office-2013-to-use-additional-languages-for-display-or-editing"></a><span data-ttu-id="9f5f2-195">Configurar o Office 2013 para usar idiomas adicionais para exibi??o ou edi??o</span><span class="sxs-lookup"><span data-stu-id="9f5f2-195">Configure Office 2013 to use additional languages for display or editing</span></span>

<span data-ttu-id="9f5f2-p124">Voc? pode usar um Pacote de idiomas do Office 2013 para instalar um idioma adicional. Para saber mais sobre os Pacotes de idioma e onde obt?-los, veja [Op??es de idioma do Office 2013](http://office.microsoft.com/en-us/language-packs/).</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p124">You can use an Office 2013 Language pack to install an additional language. For more information about Language Packs and where to get them, see [Office 2013 Language Options](http://office.microsoft.com/en-us/language-packs/).</span></span>

> <span data-ttu-id="9f5f2-p125">**OBSERVA??O:** se voc? for assinante do MSDN, ? poss?vel que j? tenha os Pacotes de Idiomas do Office 2013. Para determinar se a sua assinatura oferece Pacotes de Idiomas do Office 2013 para download, v? para [P?gina Inicial de Assinaturas do MSDN](https://msdn.microsoft.com/subscriptions/manage/), insira Pacote de Idiomas do Office 2013 em **Downloads de Softwares**, escolha **Pesquisa** e selecione **Produtos dispon?veis com minha assinatura**. Em **Idioma**, marque a caixa de sele??o do Pacote de Idiomas que voc? deseja baixar e, em seguida, selecione **Ir**.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p125">**NOTE** If you are an MSDN Subscriber, you might already have the Office 2013 Language Packs available to you. To determine whether your subscription offers Office 2013 Language Packs for download, go to [MSDN Subscriptions Home](https://msdn.microsoft.com/subscriptions/manage/), enter Office 2013 Language Pack in **Software downloads**, choose **Search**, and then select **Products available with my subscription**. Under **Language**, select the check box for the Language Pack you want to download, and then choose  **Go**.</span></span> 

<span data-ttu-id="9f5f2-p126">Depois de instalar o Pacote de Idiomas, voc? pode configurar o Office 2013 para usar o idioma instalado para exibir na interface do usu?rio, para edi??o de conte?do do documento, ou ambos. O exemplo neste artigo usa uma instala??o do Office 2013 que tenha o Pacote de Idiomas do espanhol aplicado.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p126">After you install the Language Pack, you can configure Office 2013 to use the installed language for display in the UI, for editing document content, or both. The example in this article uses an installation of Office 2013 that has the Spanish Language Pack applied.</span></span>

### <a name="create-an-office-add-in-project"></a><span data-ttu-id="9f5f2-203">Criar um projeto de Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="9f5f2-203">Create an Office Add-in project</span></span>

1. <span data-ttu-id="9f5f2-204">No Visual Studio, escolha **Arquivo**  >  **Novo Projeto**.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-204">In Visual Studio, choose **File** > **New Project**.</span></span>
    
2. <span data-ttu-id="9f5f2-205">Na caixa de di?logo **Novo Projeto**, em **Modelos**, expanda **Visual Basic** ou **Visual C#**, expanda **Office/SharePoint** e, em seguida, selecione **Suplementos do Office**.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-205">In the **New Project** dialog box, under **Templates**, expand **Visual Basic** or **Visual C#**, expand **Office/SharePoint**, and then choose  **Office Add-ins**.</span></span>
    
3. <span data-ttu-id="9f5f2-p127">Escolha **Suplemento do Office** e, em seguida, nomeie seu suplemento, por exemplo WorldReadyAddIn. Escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p127">Choose **Office Add-in**, and then name your add-in, for example WorldReadyAddIn. Choose  **OK**.</span></span>
    
4. <span data-ttu-id="9f5f2-p128">Na caixa de di?logo **Criar Suplemento do Office**, selecione **Painel de tarefas** e selecione **Pr?ximo**. Na pr?xima p?gina, desmarque e marque as caixas de todos os aplicativos, exceto do Word. Selecione **Concluir** para criar o projeto.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p128">In the **Create Office Add-in** dialog box, select **Task pane** and choose **Next**. On the next page, clear the check boxes for all host applications except Word. Choose **Finish** to create the project.</span></span>
    

### <a name="localize-the-text-used-in-your-add-in"></a><span data-ttu-id="9f5f2-211">Localizar o texto usado no seu suplemento</span><span class="sxs-lookup"><span data-stu-id="9f5f2-211">Localize the text used in your add-in</span></span>

<span data-ttu-id="9f5f2-212">O texto que voc? deseja localizar para outro idioma aparece em duas ?reas:</span><span class="sxs-lookup"><span data-stu-id="9f5f2-212">The text that you want to localize for another language appears in two areas:</span></span>

-  <span data-ttu-id="9f5f2-p129">**Nome de exibi??o e descri??o do suplemento**. Isso ? controlado por entradas no arquivo do manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p129">**Add-in display name and description**. This is controlled by entries in the add-in manifest file.</span></span>
    
-  <span data-ttu-id="9f5f2-p130">**Interface do Usu?rio do Suplemento**. Voc? pode localizar as cadeias de caracteres que aparecem na interface do usu?rio do seu suplemento usando c?digos do JavaScript, por exemplo, usando um arquivo de recurso separado que contenha as cadeias de caracteres localizadas.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p130">**Add-in UI**. You can localize the strings that appear in your add-in UI by using JavaScript code???for example, by using a separate resource file that contains the localized strings.</span></span>
    
<span data-ttu-id="9f5f2-217">Para localizar o nome de exibi??o e a descri??o do suplemento:</span><span class="sxs-lookup"><span data-stu-id="9f5f2-217">To localize the add-in display name and description:</span></span>

1. <span data-ttu-id="9f5f2-218">Em **Gerenciador de Solu??es**, expanda **WorldReadyAddIn**, **WorldReadyAddInManifest** e, em seguida, selecione **WorldReadyAddIn.xml**.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-218">In **Solution Explorer**, expand **WorldReadyAddIn**, **WorldReadyAddInManifest**, and then choose  **WorldReadyAddIn.xml**.</span></span>
    
2. <span data-ttu-id="9f5f2-219">No WorldReadyAddInManifest.xml, substitua os elementos [DisplayName] e [Description] com o seguinte bloqueio de c?digo:</span><span class="sxs-lookup"><span data-stu-id="9f5f2-219">In WorldReadyAddInManifest.xml, replace the [DisplayName] and [Description] elements with the following block of code:</span></span>
    
    > <span data-ttu-id="9f5f2-220">**OBSERVA??O:** voc? pode substituir as cadeias de caracteres do idioma espanhol localizado usadas neste exemplo pelos elementos [DisplayName] e [Description] pelas cadeias de caracteres localizadas de qualquer outro idioma.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-220">**NOTE** You can replace the Spanish language localized strings used in this example for the [DisplayName] and [Description] elements with the localized strings for any other language.</span></span>

    ```xml
    <DisplayName DefaultValue="World Ready add-in">
      <Override Locale="es-es" Value="Aplicación de uso internacional"/>
    </DisplayName>
    <Description DefaultValue="An add-in for testing localization">
      <Override Locale="es-es" Value="Una aplicación para la prueba de la localización"/>
    </Description>
    ```

3. <span data-ttu-id="9f5f2-221">Quando voc? altera o idioma de exibi??o do Office 2013 do ingl?s para o espanhol, por exemplo, e executa o suplemento, o nome de exibi??o do suplemento e a descri??o s?o mostrados com texto localizado.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-221">When you change the display language for Office 2013 from English to Spanish, for example, and then run the add-in, the add-in display name and description are shown with localized text.</span></span> 
    
<span data-ttu-id="9f5f2-222">Para definir a interface do usu?rio do suplemento:</span><span class="sxs-lookup"><span data-stu-id="9f5f2-222">To lay out the add-in UI:</span></span>

1. <span data-ttu-id="9f5f2-223">No Visual Studio, no **Gerenciador de Solu??es**, selecione **Home.html**.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-223">In Visual Studio, in **Solution Explorer**, choose **Home.html**.</span></span>
    
2. <span data-ttu-id="9f5f2-224">Substitua o HTML em Home.html pelo seguinte HTML.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-224">Replace the HTML in Home.html with the following HTML.</span></span>
    
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

3. <span data-ttu-id="9f5f2-225">No Visual Studio, selecione **Arquivo**,  **Salvar Suplemento\Home\Home.html**.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-225">In Visual Studio, choose  **File**,  **Save AddIn\Home\Home.html**.</span></span>
    
<span data-ttu-id="9f5f2-226">A figura a seguir mostra o elemento do cabe?alho (h1) e o elemento do par?grafo (p) que exibir? o texto localizado quando seu suplemento de amostra for executado.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-226">The following figure shows the heading (h1) element and the paragraph (p) element that will display localized text when your sample add-in runs.</span></span>

<span data-ttu-id="9f5f2-227">*Figura 1. A interface do usu?rio do suplemento*</span><span class="sxs-lookup"><span data-stu-id="9f5f2-227">*Figure 1. The add-in UI*</span></span>

![Interface de usu?rio do aplicativo com as se??es real?adas.](../images/office15-app-how-to-localize-fig03.png)

### <a name="add-the-resource-file-that-contains-the-localized-strings"></a><span data-ttu-id="9f5f2-229">Adicionar o arquivo de recurso que cont?m as cadeias de caracteres localizadas</span><span class="sxs-lookup"><span data-stu-id="9f5f2-229">Add the resource file that contains the localized strings</span></span>

<span data-ttu-id="9f5f2-p131">O arquivo de recurso do JavaScript cont?m as cadeias de caracteres usadas para a interface do usu?rio do suplemento. A interface do usu?rio do suplemento de amostra tem um elemento h1 que exibe uma sauda??o e um elemento p que apresenta o suplemento ao usu?rio.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p131">The JavaScript resource file contains the strings used for the add-in UI. The sample add-in UI has an h1 element that displays a greeting, and a p element that introduces the add-in to the user.</span></span> 

<span data-ttu-id="9f5f2-p132">Para habilitar cadeias de caracteres para o cabe?alho e par?grafo, coloque as cadeias de caracteres em um arquivo de recurso separado. O arquivo de recurso cria um objeto do JavaScript que cont?m um objeto JSON (JavaScript Object Notation) separado para cada conjunto de cadeias de caracteres localizadas. O arquivo de recurso tamb?m fornece um m?todo para obter o objeto JSON apropriado de volta para uma determinada localidade.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p132">To enable localized strings for the heading and paragraph, you place the strings in a separate resource file. The resource file creates a JavaScript object that contains a separate JavaScript Object Notation (JSON) object for each set of localized strings. The resource file also provides a method for getting back the appropriate JSON object for a given locale.</span></span> 

<span data-ttu-id="9f5f2-235">Para adicionar o arquivo de recurso ao projeto do suplemento:</span><span class="sxs-lookup"><span data-stu-id="9f5f2-235">To add the resource file to the add-in project:</span></span>

1. <span data-ttu-id="9f5f2-236">No **Gerenciador de Solu??es** no Visual Studio, escolha a pasta **Suplemento** no projeto da Web para o suplemento de amostra e selecione **Adicionar**  >  **Arquivo JavaScript**.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-236">In **Solution Explorer** in Visual Studio, choose the **Add-in** folder in the web project for the sample add-in, and choose **Add** > **JavaScript file**.</span></span>
    
2. <span data-ttu-id="9f5f2-237">Na caixa de di?logo **Especificar o nome do item**, insira UIStrings.js.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-237">In the **Specify Name for Item** dialog box, enterUIStrings.js.</span></span>
    
3. <span data-ttu-id="9f5f2-238">Adicione o c?digo a seguir ao arquivo UIStrings.js.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-238">Add the following code to the UIStrings.js file.</span></span>

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

<span data-ttu-id="9f5f2-239">O arquivo de recurso UIStrings.js cria o objeto, **UIStrings**, que cont?m as cadeias de caracteres localizadas para a interface do usu?rio do suplemento.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-239">The UIStrings.js resource file creates an object, **UIStrings**, which contains the localized strings for your add-in UI.</span></span> 

### <a name="localize-the-text-used-for-the-add-in-ui"></a><span data-ttu-id="9f5f2-240">Localizar o texto usado na interface do usu?rio do suplemento</span><span class="sxs-lookup"><span data-stu-id="9f5f2-240">Localize the text used for the add-in UI</span></span>

<span data-ttu-id="9f5f2-p133">Para usar o arquivo de recurso no seu suplemento, voc? precisar? adicionar a ele uma marca de script em Home.html. Quando Home.html for carregado, o UIStrings.js ser? executado e o objeto **UIStrings** que voc? utiliza para obter a cadeia de caracteres ficar? dispon?vel para seu c?digo. Adicione o seguinte HTML ? marca de cabe?alho do Home.html para tornar **UIStrings** dispon?vel para seu c?digo.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p133">To use the resource file in your add-in, you'll need to add a script tag for it on Home.html. When Home.html is loaded, UIStrings.js executes and the **UIStrings** object that you use to get the strings is available to your code. Add the following HTML in the head tag for Home.html to make **UIStrings** available to your code.</span></span>

```html
<!-- Resource file for localized strings:                                                          -->
<script src="../UIStrings.js" type="text/javascript"></script>
```

<span data-ttu-id="9f5f2-244">Agora voc? pode usar o objeto **UIStrings** para definir as cadeias de caracteres da interface do usu?rio do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-244">Now you can use the **UIStrings** object to set the strings for the UI of your add-in.</span></span>

<span data-ttu-id="9f5f2-p134">Se voc? quiser alterar a localiza??o do seu suplemento com base no idioma usado para exibi??o nos menus e comandos no aplicativo host, use a propriedade **Office.context.displayLanguage** para obter a localidade desse idioma. Por exemplo, se o idioma do aplicativo host utilizar espanhol para exibir menus e comandos, a propriedade **Office.context.displayLanguage** retornar? o c?digo es-ES.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p134">If you want to change the localization for your add-in based on what language is used for display in menus and commands in the host application, you use the **Office.context.displayLanguage** property to get the locale for that language. For example, if the host application language uses Spanish for display in menus and commands, the **Office.context.displayLanguage** property will return the language code es-ES.</span></span>

<span data-ttu-id="9f5f2-p135">Se voc? quiser alterar a localiza??o do seu suplemento com base no idioma que est? sendo usado para editar o conte?do do documento, use a propriedade **Office.context.contentLanguage** para obter a localidade do idioma. Por exemplo, se o idioma do aplicativo host utilizar espanhol para editar o conte?do do documento, a propriedade **Office.context.contentLanguage** retornar? o c?digo es-ES.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p135">If you want to change the localization for your add-in based on what language is being used for editing document content, you use the  **Office.context.contentLanguage** property to get the locale for that language. For example, if the host application language uses Spanish for editing document content, the **Office.context.contentLanguage** property will return the language code es-ES.</span></span>

<span data-ttu-id="9f5f2-249">Depois que voc? souber o idioma que o aplicativo host est? utilizando, ? poss?vel usar **UIStrings** para obter o conjunto de cadeias de caracteres localizadas correspondentes ao idioma do aplicativo host.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-249">After you know the language the host application is using, you can use **UIStrings** to get the set of localized strings that matches the host application language.</span></span>

<span data-ttu-id="9f5f2-p136">Substitua o c?digo no arquivo Home.js pelo c?digo a seguir. O c?digo mostra como voc? pode alterar as cadeias de caracteres usadas nos elementos da interface do usu?rio no Home.html com base no idioma de exibi??o do aplicativo host ou no idioma de edi??o do aplicativo host.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p136">Replace the code in the Home.js file with the following code. The code shows how you can change the strings used in the UI elements on Home.html based on either the display language of the host application or the editing language of the host application.</span></span>

> <span data-ttu-id="9f5f2-252">**OBSERVA??O:** para alternar entre a altera??o da localiza??o do suplemento com base no idioma usado para edi??o, remova o coment?rio da linha de c?digo `var myLanguage = Office.context.contentLanguage;` e inclua o coment?rio na linha de c?digo `var myLanguage = Office.context.displayLanguage;`</span><span class="sxs-lookup"><span data-stu-id="9f5f2-252">**NOTE** To switch between changing the localization of the add-in based on the language used for editing, uncomment the line of code  `var myLanguage = Office.context.contentLanguage;` and comment out the line of code `var myLanguage = Office.context.displayLanguage;`</span></span>

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

### <a name="test-your-localized-add-in"></a><span data-ttu-id="9f5f2-253">Testar seu suplemento localizado</span><span class="sxs-lookup"><span data-stu-id="9f5f2-253">Test your localized add-in</span></span>

<span data-ttu-id="9f5f2-254">Para testar seu suplemento localizado, altere o idioma usado para exibir ou editar no aplicativo host e execute o seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-254">To test your localized add-in, change the language used for display or editing in the host application and then run your add-in.</span></span> 

<span data-ttu-id="9f5f2-255">Para alterar o idioma usado para exibir ou editar no seu suplemento:</span><span class="sxs-lookup"><span data-stu-id="9f5f2-255">To change the language used for display or editing in your add-in:</span></span>

1. <span data-ttu-id="9f5f2-p137">No Word 2013, selecione **Arquivo** > , **Op??es** > , **Idioma**. A figura a seguir mostra a caixa de di?logo **Op??es do Word** aberta na guia Idioma.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p137">In Word 2013, choose **File** > **Options** > **Language**. The following figure shows the **Word Options** dialog box opened to the Language tab.</span></span>
    
    <span data-ttu-id="9f5f2-258">*Figura 2. Op??es de idioma na caixa de di?logo Op??es do Word 2013*</span><span class="sxs-lookup"><span data-stu-id="9f5f2-258">*Figure 2. Language options in the Word 2013 Options dialog box*</span></span>

    ![Caixa de di?logo Op??es do Word 2013.](../images/office15-app-how-to-localize-fig04.png)

2. <span data-ttu-id="9f5f2-p138">Em **Escolher idiomas de exibi??o e da ajuda**, selecione o idioma desejado para exibi??o, por exemplo, espanhol, e selecione a seta para cima para mover o idioma espanhol para a primeira posi??o na lista. Ou, para alterar o idioma usado para edi??o, em **Escolher idiomas de edi??o**, escolha o idioma que voc? deseja usar para edi??o, por exemplo, espanhol, e selecione **Definir como Padr?o**.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p138">Under **Choose Display and Help Languages**, select the language that you want for display, for example Spanish, and then choose the up arrow to move the Spanish language to the first position in the list. Alternatively, to change the language used for editing, under  **Choose editing languages**, choose the language you want to use for editing, for example, Spanish, and then choose **Set as Default**.</span></span>
    
3. <span data-ttu-id="9f5f2-262">Escolha **OK** para confirmar sua sele??o e feche o Word.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-262">Choose **OK** to confirm your selection, and then close Word.</span></span>
    
<span data-ttu-id="9f5f2-p139">Execute o suplemento de exemplo. O suplemento do painel de tarefas ? carregado no Word 2013 e as cadeias de caracteres na interface do usu?rio do suplemento s?o alteradas para corresponder ao idioma usado pelo aplicativo host, conforme mostrado na figura a seguir.</span><span class="sxs-lookup"><span data-stu-id="9f5f2-p139">Run the sample add-in. The taskpane add-in loads in Word 2013, and the strings in the add-in UI change to match the language used by the host application, as shown in the following figure.</span></span>


<span data-ttu-id="9f5f2-265">*Figura 3. Interface do usu?rio do suplemento com o texto localizado*</span><span class="sxs-lookup"><span data-stu-id="9f5f2-265">*Figure 3. Add-in UI with localized text*</span></span>

![Aplicativo com texto localizado da interface do usu?rio.](../images/office15-app-how-to-localize-fig05.png)

## <a name="see-also"></a><span data-ttu-id="9f5f2-267">Confira tamb?m</span><span class="sxs-lookup"><span data-stu-id="9f5f2-267">See also</span></span>

- [<span data-ttu-id="9f5f2-268">Diretrizes de design para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="9f5f2-268">Design guidelines for Office Add-ins</span></span>](../design/add-in-design.md)    
- [<span data-ttu-id="9f5f2-269">Identificadores de idioma e valores da ID de OptionState no Office 2013</span><span class="sxs-lookup"><span data-stu-id="9f5f2-269">Language identifiers and OptionState Id values in Office 2013</span></span>](http://technet.microsoft.com/en-us/library/cc179219%28Office.15%29.aspx)

[DefaultLocale]:        https://dev.office.com/reference/add-ins/manifest/defaultlocale
[Descri??o]:          https://dev.office.com/reference/add-ins/manifest/description
[Description]:          https://dev.office.com/reference/add-ins/manifest/description
[DisplayName]:          https://dev.office.com/reference/add-ins/manifest/displayname
[IconUrl]:              https://dev.office.com/reference/add-ins/manifest/iconurl
[HighResolutionIconUrl]:https://dev.office.com/reference/add-ins/manifest/highresolutioniconurl
[Recursos]:            https://dev.office.com/reference/add-ins/manifest/resources
[Resources]:            https://dev.office.com/reference/add-ins/manifest/resources
[SourceLocation]:       https://dev.office.com/reference/add-ins/manifest/sourcelocation
[Substitui??o]:             https://dev.office.com/reference/add-ins/manifest/override
[Override]:             https://dev.office.com/reference/add-ins/manifest/override
[DesktopSettings]:      https://dev.office.com/reference/add-ins/manifest/desktopsettings
[TabletSettings]:       https://dev.office.com/reference/add-ins/manifest/tabletsettings
[PhoneSettings]:        https://dev.office.com/reference/add-ins/manifest/phonesettings
[displayLanguage]:  https://dev.office.com/reference/add-ins/shared/office.context.displaylanguage 
[contentLanguage]:  https://dev.office.com/reference/add-ins/shared/office.context.contentlanguage 
[RFC 3066]: https://www.rfc-editor.org/info/rfc3066
