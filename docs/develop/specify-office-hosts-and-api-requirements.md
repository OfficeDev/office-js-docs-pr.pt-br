---
title: Especificar hosts do Office e requisitos de API
description: Saiba como especificar aplicativos do Office e requisitos de API para que seu complemento funcione conforme o esperado.
ms.date: 08/24/2020
localization_priority: Normal
ms.openlocfilehash: 948e86e99150ebf2d0bc7deaa5512627679b025f
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237837"
---
# <a name="specify-office-applications-and-api-requirements"></a>Especificar requisitos da API e de aplicativos do Office

Seu Complemento do Office pode depender de um aplicativo específico do Office, um conjunto de requisitos, um membro da API ou uma versão da API para funcionar conforme o esperado. Por exemplo, o suplemento pode:

- Executar em um único aplicativo do Office (por exemplo, Word ou Excel) ou diversos aplicativos.

- Usar as APIs de JavaScript que estão disponíveis apenas em algumas versões do Office. Por exemplo, você pode usar as APIs JavaScript do Excel em um suplemento executado no Excel 2016.

- Executar apenas nas versões do Office que oferecem suporte a membros da API que seu suplemento usa.

Este artigo ajuda você a entender quais opções você deve escolher para garantir que seu suplemento funcione conforme o esperado e atinja o público mais amplo possível.

> [!NOTE]
> Para uma visão de alto nível de onde os Complementos do Office têm suporte no momento, confira a página disponibilidade de aplicativos e plataformas do cliente Office para Os [Complementos do Office.](../overview/office-add-in-availability.md)

A tabela a seguir lista os principais conceitos discutidos neste artigo.

|**Conceito**|**Descrição**|
|:-----|:-----|
|Aplicativo do Office, aplicativo cliente do Office|O aplicativo do Office usado para executar seu suplemento. Por exemplo, Word e assim por diante.|
|Plataforma|Onde o aplicativo do Office é executado, como em um navegador ou em um iPad.|
|Conjunto de requisitos|Um grupo nomeado de membros relacionados da API. Os complementos usam conjuntos de requisitos para determinar se o aplicativo do Office dá suporte a membros da API usados pelo seu complemento. É mais fácil testar se há suporte para um conjunto de requisitos do que o suporte para membros individuais da API. O suporte ao conjunto de requisitos varia de acordo com o aplicativo do Office e a versão do aplicativo do Office. <br >Conjuntos de requisitos são especificados no arquivo de manifesto. Ao especificar conjuntos de requisitos no manifesto, você define o nível mínimo de suporte à API que o aplicativo do Office deve fornecer para executar o seu complemento. Os aplicativos do Office que não suportam conjuntos de <span class="ui">requisitos especificados</span>no manifesto não podem executar o seu complemento e o seu complemento não será exibido em Meus Complementos. Isso restringe onde o seu complemento está disponível. No código usando verificações de tempo de execução. Para obter uma lista completa de conjuntos de requisitos, confira [Conjuntos de requisitos de Suplemento do Office](../reference/requirement-sets/office-add-in-requirement-sets.md).|
|Verificação no tempo de execução|Um teste executado em tempo de execução para determinar se o aplicativo do Office que está executando o seu add-in oferece suporte a conjuntos de requisitos ou métodos usados pelo seu complemento. Para executar uma verificação de tempo de execução, use uma instrução **if** com o método, os conjuntos de requisitos ou os nomes de método que não fazem parte de um `isSetSupported` conjunto de requisitos. Use as verificações no tempo de execução para garantir que seu suplemento alcance o maior número de clientes. Ao contrário dos conjuntos de requisitos, as verificações de tempo de execução não especificam o nível mínimo de suporte à API que o aplicativo do Office deve fornecer para que o seu complemento seja executado. Em vez disso, use a **instrução if** para determinar se um membro da API é suportado. Se houver, você poderá proporcionar mais funcionalidade em seu suplemento. Seu suplemento sempre aparecerá em **Meus Suplementos** ao usar verificações no tempo de execução.|

## <a name="before-you-begin"></a>Antes de começar

O suplemento deve usar a versão mais recente do esquema de manifesto de suplemento. Se você usar verificações de tempo de execução no seu complemento, certifique-se de usar a biblioteca mais recente da API JavaScript do Office (office.js).

### <a name="specify-the-latest-add-in-manifest-schema"></a>Especificar o esquema de manifesto de suplemento mais recente

Seu manifesto de suplemento deve usar a versão 1.1 do esquema de manifesto de suplemento. De `OfficeApp` definir o elemento no manifesto do seu complemento da seguinte forma.

```XML
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
```

### <a name="specify-the-latest-office-javascript-api-library"></a>Especificar a biblioteca mais recente da API JavaScript do Office

Se você usar verificações de tempo de execução, consulte a versão mais recente da biblioteca da API JavaScript do Office da CDN (rede de distribuição de conteúdo). Para tanto, adicione a seguinte marca `script` ao código HTML. O uso de `/1/` na URL da CDN garante a referência à versão mais recente do Office.js.

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## <a name="options-to-specify-office-applications-or-api-requirements"></a>Opções para especificar aplicativos do Office ou requisitos de API

Quando você especifica aplicativos do Office ou requisitos de API, há vários fatores a considerar. O diagrama a seguir mostra como decidir sobre qual técnica usar em seu suplemento.

![Escolha a melhor opção para o seu complemento ao especificar aplicativos do Office ou requisitos de API](../images/options-for-office-hosts.png)

- Se o seu complemento for executado em um aplicativo do Office, de definir `Hosts` o elemento no manifesto. Para saber mais, confira [Definir o elemento Hosts](#set-the-hosts-element).

- Para definir o conjunto de requisitos mínimos ou membros da API que um aplicativo do Office deve suportar para executar o seu complemento, de definir `Requirements` o elemento no manifesto. Para saber mais, confira [Definir o elemento Requirements no manifesto](#set-the-requirements-element-in-the-manifest).

- Se você quiser fornecer funcionalidade adicional se conjuntos de requisitos específicos ou membros da API estão disponíveis no aplicativo do Office, execute uma verificação de tempo de execução no código JavaScript do seu complemento. Por exemplo, se o suplemento for executado no Excel 2016, use os membros do API JavaScript do Excel a fim de fornecer funcionalidades adicionais. Para saber mais, confira [Usar verificações de tempo de execução em seu código JavaScript](#use-runtime-checks-in-your-javascript-code).

## <a name="set-the-hosts-element"></a>Definir o elemento Hosts

Para fazer com que o seu complemento seja executado em um aplicativo cliente do Office, use os elementos `Hosts` `Host` no manifesto. Se você não especificar o elemento, o seu complemento será executado em todos os aplicativos do Office com suporte dos `Hosts` Complementos do Office.

Por exemplo, o exemplo a seguir e a declaração especifica que o complemento funcionará com qualquer versão do Excel, que inclui o Excel na Web, o Windows e `Hosts` `Host` o iPad.

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

O `Hosts` elemento pode conter um ou mais `Host` elementos. O `Host` elemento especifica o aplicativo do Office que seu add-in requer. O `Name` atributo é obrigatório e pode ser definido como um dos seguintes valores.

| Nome          | Aplicativos cliente do Office                      |
|:--------------|:----------------------------------------------|
| Banco de dados      | Aplicativos Web do Access                               |
| Documento      | Word na Web, Windows, Mac, iPad           |
| Caixa de correio       | Outlook na Web, Windows, Mac, Android, iOS|
| Presentation  | PowerPoint na Web, Windows, Mac, iPad     |
| Project       | Project no Windows                            |
| Pasta de Trabalho      | Excel na Web, Windows, Mac, iPad          |

> [!NOTE]
> O `Name` atributo especifica o aplicativo cliente do Office que pode executar o seu complemento. Os aplicativos do Office têm suporte em diferentes plataformas e são executados em desktops, navegadores da Web, tablets e dispositivos móveis. Você não pode especificar qual plataforma pode ser usada para executar seu suplemento. Por exemplo, se você especificar , tanto o Outlook na Web quanto o Windows podem ser usados `Mailbox` para executar o seu complemento.

> [!IMPORTANT]
> Não recomendamos mais criar e usar aplicativos Web do Access e bancos de dados no SharePoint. Como alternativa, use o [Microsoft PowerApps](https://powerapps.microsoft.com/) para criar soluções de negócios sem código para dispositivos móveis e Web.

## <a name="set-the-requirements-element-in-the-manifest"></a>Definir o elemento Requirements no manifesto

O elemento especifica os conjuntos de requisitos mínimos ou membros da API que devem ter suporte do aplicativo do Office para executar `Requirements` o seu complemento. O `Requirements` elemento pode especificar conjuntos de requisitos e métodos individuais usados no seu complemento. Na versão 1.1 do esquema de manifesto do complemento, o elemento é opcional para todos os complementos, exceto para os do `Requirements` Outlook.

> [!WARNING]
> Use o elemento somente para especificar conjuntos de requisitos críticos ou membros da `Requirements` API que seu complemento deve usar. Se o aplicativo ou a plataforma do Office não oferece suporte aos conjuntos de requisitos ou membros da API especificados no elemento, o complemento não será executado nesse aplicativo ou plataforma e não será exibido em Meus `Requirements` **Complementos.** Em vez disso, recomendamos disponibilizar seu complemento em todas as plataformas de um aplicativo do Office, como o Excel na Web, o Windows e o iPad. Para disponibilizar seu complemento em todos os aplicativos e plataformas do  _Office,_ use verificações de tempo de execução em vez do `Requirements` elemento.

O exemplo de código a seguir mostra um complemento que é carregado em todos os aplicativos cliente do Office que suportam o seguinte:

-  `TableBindings` conjunto de requisitos, que tem uma versão mínima de "1.1".

-  `OOXML` conjunto de requisitos, que tem uma versão mínima de "1.1".

-  `Document.getSelectedDataAsync` .

```XML
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="TableBindings" MinVersion="1.1"/>
      <Set Name="OOXML" MinVersion="1.1"/>
   </Sets>
   <Methods>
      <Method Name="Document.getSelectedDataAsync"/>
   </Methods>
</Requirements>
```

- O `Requirements` elemento contém os elementos e `Sets` `Methods` filho.

- O `Sets` elemento pode conter um ou mais `Set` elementos. `DefaultMinVersion` especifica o valor padrão `MinVersion` de todos os elementos `Set` filho.

- O `Set` elemento especifica conjuntos de requisitos que o aplicativo do Office deve suportar para executar o complemento. O `Name` atributo especifica o nome do conjunto de requisitos. Especifica `MinVersion` a versão mínima do conjunto de requisitos. `MinVersion` substitui o valor de For `DefaultMinVersion` more information about requirement sets and requirement set versions that your API members belong to, see [Office Add-in requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).

- O `Methods` elemento pode conter um ou mais `Method` elementos. Você não pode usar o `Methods` elemento com os complementos do Outlook.

- O elemento especifica um método individual que deve ter suporte `Method` no aplicativo do Office em que o seu complemento é executado. O `Name` atributo é obrigatório e especifica o nome do método qualificado com seu objeto pai.

## <a name="use-runtime-checks-in-your-javascript-code"></a>Usar verificações no tempo de execução em seu código JavaScript

Talvez você queira fornecer funcionalidade adicional no seu complemento se determinados conjuntos de requisitos são suportados pelo aplicativo do Office. Por exemplo, você pode usar a nova API JavaScript do Word em seu suplemento existente se o suplemento for executado no Word 2016.  Para fazer isso, use o método [isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) com o nome do conjunto de requisitos. `isSetSupported` determina, em tempo de execução, se o aplicativo do Office que está executando o complemento dá suporte ao conjunto de requisitos. Se o conjunto de requisitos for suportado, retornará true e executa o código adicional que usa os membros `isSetSupported` da API desse conjunto de requisitos.  Se o aplicativo do Office não suportar o conjunto de requisitos, retornará false e `isSetSupported` o código adicional não será executado.  O código a seguir mostra a sintaxe a ser usada com `isSetSupported` .

```js
if (Office.context.requirements.isSetSupported(RequirementSetName, MinimumVersion))
{
   // Code that uses API members from RequirementSetName.
}

```

- _RequirementSetName_ (obrigatório) é uma cadeia de caracteres que representa o nome do conjunto de requisitos (por exemplo, "**ExcelApi**", "**Mailbox**", etc.). Para saber mais sobre os conjuntos de requisitos disponíveis, confira [Conjuntos de requisitos de Suplemento do Office](../reference/requirement-sets/office-add-in-requirement-sets.md).
- _MinimumVersion_ (opcional) é uma cadeia de caracteres que especifica a versão mínima do conjunto de requisitos que o aplicativo do Office deve suportar para que o código dentro da instrução seja executado `if` (por exemplo, "**1.9**").

> [!WARNING]
> Ao chamar o `isSetSupported` método, o valor do `MinimumVersion` parâmetro (se especificado) deve ser uma cadeia de caracteres. Isso ocorre porque o analisador de JavaScript não pode diferenciar valores numéricos, como 1.1 e 1.10, onde é possível para valores de cadeia de caracteres como "1.1" e "1.10".
> A sobrecarga de `number` está obsoleta.

Use `isSetSupported` com o aplicativo do Office associado da seguinte `RequirementSetName` forma.

|Aplicativo do Office|RequirementSetName|
|---|---|
|Excel|ExcelApi|
|OneNote|OneNoteApi|
|Outlook|Caixa de correio|
|Word|WordApi|

O método e os conjuntos de requisitos para esses aplicativos estão disponíveis no arquivo `isSetSupported` de Office.js mais recente na CDN. Se você não usar o Office.js cdn, seu complemento pode gerar exceções porque `isSetSupported` será indefinido. Para saber mais, confira [Especificar a biblioteca mais recente da API JavaScript do Office.](#specify-the-latest-office-javascript-api-library)

O exemplo de código a seguir mostra como um complemento pode fornecer funcionalidades diferentes para diferentes aplicativos do Office que podem dar suporte a diferentes conjuntos de requisitos ou membros da API.

```js
if (Office.context.requirements.isSetSupported('WordApi', '1.1'))
{
    // Run code that provides additional functionality using the Word JavaScript API when the add-in runs in Word 2016 or later.
}
else if (Office.context.requirements.isSetSupported('CustomXmlParts'))
{
    // Run code that uses API members from the CustomXmlParts requirement set.
}
else
{
    // Run additional code when the Office application is not Word 2016 or later and does not support the CustomXmlParts requirement set.
}

```

## <a name="runtime-checks-using-methods-not-in-a-requirement-set"></a>Verificações no tempo de execução usando métodos que não fazem parte de um conjunto de requisitos

Alguns membros de API não pertencem a conjuntos de requisitos. Isso só se aplica a membros da API que fazem parte do namespace da [API JavaScript](../reference/javascript-api-for-office.md) do Office (qualquer coisa em, exceto APIs de Caixa de Correio do Outlook), mas não a membros da API que pertencem à API JavaScript do Word (qualquer coisa em ), API JavaScript do Excel (qualquer coisa em ) ou `Office.` API JavaScript do [](/javascript/api/outlook)OneNote (qualquer coisa em ) [](../reference/overview/word-add-ins-reference-overview.md) `Word.` [](../reference/overview/excel-add-ins-reference-overview.md) `Excel.` [](../reference/overview/onenote-add-ins-javascript-reference.md) `OneNote.` namespaces. Quando o seu complemento depende de um método que não faz parte de um conjunto de requisitos, você pode usar a verificação de tempo de execução para determinar se o método é suportado pelo aplicativo do Office, conforme mostrado no exemplo de código a seguir. Para obter uma lista completa dos métodos que não pertencem a um conjunto de requisitos, confira [Conjuntos de requisitos de Suplemento do Office](../reference/requirement-sets/office-add-in-requirement-sets.md#methods-that-arent-part-of-a-requirement-set).

> [!NOTE]
> Recomendamos limitar o uso desse tipo de verificação no tempo de execução no código de seu suplemento.

O exemplo de código a seguir verifica se o aplicativo do Office oferece `document.setSelectedDataAsync` suporte.

```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses `document.setSelectedDataAsync`.
}
```


## <a name="see-also"></a>Confira também

- [Manifesto XML dos Suplementos do Office](add-in-manifests.md)
- [Conjuntos de requisitos de Suplemento do Office](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)
