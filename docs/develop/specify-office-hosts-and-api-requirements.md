---
title: Especificar hosts do Office e requisitos de API
description: Saiba como especificar aplicativos do Office e requisitos de API para que o suplemento funcione conforme o esperado.
ms.date: 08/24/2020
localization_priority: Normal
ms.openlocfilehash: 90ee7c3a5ad01252336608c02f995bbcbbe94212
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292626"
---
# <a name="specify-office-applications-and-api-requirements"></a>Especificar aplicativos do Office e requisitos de API

O suplemento do Office pode depender de um aplicativo específico do Office, de um conjunto de requisitos, de um membro da API ou de uma versão da API para funcionar conforme o esperado. Por exemplo, o suplemento pode:

- Executar em um único aplicativo do Office (por exemplo, Word ou Excel) ou diversos aplicativos.

- Usar as APIs de JavaScript que estão disponíveis apenas em algumas versões do Office. Por exemplo, você pode usar as APIs JavaScript do Excel em um suplemento executado no Excel 2016.

- Executar apenas nas versões do Office que oferecem suporte a membros da API que seu suplemento usa.

Este artigo ajuda você a entender quais opções você deve escolher para garantir que seu suplemento funcione conforme o esperado e atinja o público mais amplo possível.

> [!NOTE]
> Para ver uma visão de alto nível do local em que os suplementos do Office têm suporte no momento, confira a página de [disponibilidade de aplicativos e suplementos da plataforma para](../overview/office-add-in-availability.md) o Office.

A tabela a seguir lista os principais conceitos discutidos neste artigo.

|**Conceito**|**Descrição**|
|:-----|:-----|
|Aplicativo do Office, aplicativo cliente do Office|O aplicativo do Office usado para executar seu suplemento. Por exemplo, Word e assim por diante.|
|Plataforma|Onde o aplicativo do Office é executado, como em um navegador ou em um iPad.|
|Conjunto de requisitos|Um grupo nomeado de membros relacionados da API. Os suplementos usam conjuntos de requisitos para determinar se o aplicativo do Office oferece suporte a membros da API usados por seu suplemento. É mais fácil testar se há suporte para um conjunto de requisitos do que o suporte para membros individuais da API. O suporte ao conjunto de requisitos varia de acordo com o aplicativo do Office e a versão do aplicativo do Office. <br >Conjuntos de requisitos são especificados no arquivo de manifesto. Ao especificar conjuntos de requisitos no manifesto, você define o nível mínimo de suporte à API que o aplicativo do Office deve fornecer para executar o suplemento. Os aplicativos do Office que não dão suporte a conjuntos de requisitos especificados no manifesto não podem executar o suplemento, e seu suplemento não será exibido em <span class="ui">meus</span>suplementos. Isso restringe o local em que o suplemento está disponível. No código usando verificações de tempo de execução. Para obter uma lista completa de conjuntos de requisitos, confira [Conjuntos de requisitos de Suplemento do Office](../reference/requirement-sets/office-add-in-requirement-sets.md).|
|Verificação no tempo de execução|Um teste executado no tempo de execução para determinar se o aplicativo do Office que está executando seu suplemento oferece suporte a conjuntos de requisitos ou métodos usados por seu suplemento. Para executar uma verificação de tempo de execução, use uma instrução **If** com o `isSetSupported` método, os conjuntos de requisitos ou os nomes dos métodos que não fazem parte de um conjunto de requisitos. Use as verificações no tempo de execução para garantir que seu suplemento alcance o maior número de clientes. Diferentemente dos conjuntos de requisitos, as verificações de tempo de execução não especificam o nível mínimo de suporte à API que o aplicativo do Office deve fornecer para que o suplemento seja executado. Em vez disso, use a instrução **If** para determinar se há suporte para um membro da API. Se houver, você poderá proporcionar mais funcionalidade em seu suplemento. Seu suplemento sempre aparecerá em **Meus Suplementos** ao usar verificações no tempo de execução.|

## <a name="before-you-begin"></a>Antes de começar

O suplemento deve usar a versão mais recente do esquema de manifesto de suplemento. Se você usar verificações de tempo de execução no seu suplemento, certifique-se de usar a biblioteca de API JavaScript do Office mais recente (office.js).

### <a name="specify-the-latest-add-in-manifest-schema"></a>Especificar o esquema de manifesto de suplemento mais recente

Seu manifesto de suplemento deve usar a versão 1.1 do esquema de manifesto de suplemento. Defina o `OfficeApp` elemento no manifesto do suplemento da seguinte maneira.

```XML
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
```

### <a name="specify-the-latest-office-javascript-api-library"></a>Especificar a biblioteca de API JavaScript do Office mais recente

Se você usar verificações de tempo de execução, faça referência à versão mais recente da biblioteca da API JavaScript do Office na CDN (rede de distribuição de conteúdo). Para tanto, adicione a seguinte marca `script` ao código HTML. O uso de `/1/` na URL da CDN garante a referência à versão mais recente do Office.js.

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## <a name="options-to-specify-office-applications-or-api-requirements"></a>Opções para especificar aplicativos do Office ou requisitos de API

Quando você especifica aplicativos ou requisitos de API do Office, há vários fatores a considerar. O diagrama a seguir mostra como decidir sobre qual técnica usar em seu suplemento.

![Escolha a melhor opção para o seu suplemento ao especificar aplicativos do Office ou requisitos de API](../images/options-for-office-hosts.png)

- Se o suplemento for executado em um aplicativo do Office, defina o `Hosts` elemento no manifesto. Para saber mais, confira [Definir o elemento Hosts](#set-the-hosts-element).

- Para definir o conjunto de requisitos mínimo ou membros da API que um aplicativo do Office deve suportar para executar seu suplemento, defina o `Requirements` elemento no manifesto. Para saber mais, confira [Definir o elemento Requirements no manifesto](#set-the-requirements-element-in-the-manifest).

- Se você quiser fornecer funcionalidade adicional se os conjuntos de requisitos específicos ou membros de API estiverem disponíveis no aplicativo do Office, execute uma verificação de tempo de execução no código JavaScript do seu suplemento. Por exemplo, se o suplemento for executado no Excel 2016, use os membros do API JavaScript do Excel a fim de fornecer funcionalidades adicionais. Para saber mais, confira [Usar verificações de tempo de execução em seu código JavaScript](#use-runtime-checks-in-your-javascript-code).

## <a name="set-the-hosts-element"></a>Definir o elemento Hosts

Para fazer com que seu suplemento seja executado em um aplicativo cliente do Office, use os `Hosts` `Host` elementos e no manifesto. Se você não especificar o `Hosts` elemento, seu suplemento será executado em todos os aplicativos do Office suportados por suplementos do Office.

Por exemplo, a `Hosts` declaração e a seguir `Host` especifica que o suplemento funcionará com qualquer versão do Excel, o que inclui o Excel na Web, Windows e iPad.

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

O `Hosts` elemento pode conter um ou mais `Host` elementos. O `Host` elemento Especifica o aplicativo do Office que seu suplemento requer. O `Name` atributo é obrigatório e pode ser definido como um dos valores a seguir.

| Nome          | Aplicativos cliente do Office                      |
|:--------------|:----------------------------------------------|
| Banco de dados      | Aplicativos Web do Access                               |
| Documento      | Word na Web, Windows, Mac, iPad           |
| Caixa de correio       | Outlook na Web, Windows, Mac, Android, iOS|
| Apresentação  | PowerPoint na Web, Windows, Mac, iPad     |
| Project       | Project no Windows                            |
| Pasta de Trabalho      | Excel na Web, Windows, Mac, iPad          |

> [!NOTE]
> O `Name` atributo especifica o aplicativo cliente do Office que pode executar seu suplemento. Os aplicativos do Office têm suporte em diferentes plataformas e são executados em desktops, navegadores da Web, tablets e dispositivos móveis. Você não pode especificar qual plataforma pode ser usada para executar seu suplemento. Por exemplo, se você especificar `Mailbox` , o Outlook na Web e o Windows podem ser usados para executar o suplemento.

> [!IMPORTANT]
> Não recomendamos mais criar e usar aplicativos Web do Access e bancos de dados no SharePoint. Como alternativa, use o [Microsoft PowerApps](https://powerapps.microsoft.com/) para criar soluções de negócios sem código para dispositivos móveis e Web.

## <a name="set-the-requirements-element-in-the-manifest"></a>Definir o elemento Requirements no manifesto

O `Requirements` elemento Especifica os conjuntos de requisitos mínimos ou membros da API que devem ser suportados pelo aplicativo do Office para executar seu suplemento. O `Requirements` elemento pode especificar os conjuntos de requisitos e os métodos individuais usados no suplemento. Na versão 1,1 do esquema de manifesto de suplemento, o `Requirements` elemento é opcional para todos os suplementos, exceto para os suplementos do Outlook.

> [!WARNING]
> Use o `Requirements` elemento para especificar conjuntos de requisitos críticos ou membros da API que o suplemento deve usar. Se o aplicativo ou a plataforma do Office não oferecer suporte aos conjuntos de requisitos ou membros de API especificados no `Requirements` elemento, o suplemento não será executado nesse aplicativo ou plataforma e não será exibido em **meus**suplementos. Em vez disso, recomendamos que você faça seu suplemento disponível em todas as plataformas de um aplicativo do Office, como Excel na Web, Windows e iPad. Para disponibilizar seu suplemento em  _todos os_ aplicativos e plataformas do Office, use as verificações de tempo de execução em vez do `Requirements` elemento.

O exemplo de código a seguir mostra um suplemento que é carregado em todos os aplicativos cliente do Office que oferecem suporte ao seguinte:

-  `TableBindings` conjunto de requisitos, que tem uma versão mínima de "1,1".

-  `OOXML` conjunto de requisitos, que tem uma versão mínima de "1,1".

-  `Document.getSelectedDataAsync` IME.

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

- O `Requirements` elemento contém os `Sets` `Methods` elementos filho e.

- O `Sets` elemento pode conter um ou mais `Set` elementos. `DefaultMinVersion` Especifica o `MinVersion` valor padrão de todos os `Set` elementos filhos.

- O `Set` elemento especifica conjuntos de requisitos que o aplicativo do Office deve suportar para executar o suplemento. O `Name` atributo especifica o nome do conjunto de requisitos. O `MinVersion` especifica a versão mínima do conjunto de requisitos. `MinVersion` Substitui o valor de `DefaultMinVersion` para obter mais informações sobre conjuntos de requisitos e versões de conjunto de requisitos aos quais seus membros da API pertencem, confira [conjuntos de requisitos de suplemento do Office](../reference/requirement-sets/office-add-in-requirement-sets.md).

- O `Methods` elemento pode conter um ou mais `Method` elementos. Você não pode usar o `Methods` elemento com os suplementos do Outlook.

- O `Method` elemento especifica um método individual que deve ser suportado no aplicativo do Office onde o suplemento é executado. O `Name` atributo é obrigatório e especifica o nome do método qualificado com seu objeto pai.

## <a name="use-runtime-checks-in-your-javascript-code"></a>Usar verificações no tempo de execução em seu código JavaScript

Você pode querer fornecer funcionalidade adicional no seu suplemento se determinados conjuntos de requisitos são compatíveis com o aplicativo do Office. Por exemplo, você pode usar a nova API JavaScript do Word em seu suplemento existente se o suplemento for executado no Word 2016.  Para fazer isso, use o método [isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) com o nome do conjunto de requisitos. `isSetSupported` determina, no tempo de execução, se o aplicativo do Office que está executando o suplemento oferece suporte ao conjunto de requisitos. Se houver suporte para o conjunto de requisitos, `isSetSupported` retornará **true** e executará o código adicional que usa os membros da API desse conjunto de requisitos. Se o aplicativo do Office não oferecer suporte ao conjunto de requisitos, `isSetSupported` retornará **false** e o código adicional não será executado. O código a seguir mostra a sintaxe a ser usada com o `isSetSupported` .

```js
if (Office.context.requirements.isSetSupported(RequirementSetName, MinimumVersion))
{
   // Code that uses API members from RequirementSetName.
}

```

- _RequirementSetName_ (obrigatório) é uma cadeia de caracteres que representa o nome do conjunto de requisitos (por exemplo, "**ExcelApi**", "**Mailbox**", etc.). Para saber mais sobre os conjuntos de requisitos disponíveis, confira [Conjuntos de requisitos de Suplemento do Office](../reference/requirement-sets/office-add-in-requirement-sets.md).
- _MinimumVersion_ (opcional) é uma cadeia de caracteres que especifica a versão do conjunto de requisitos mínimos que o aplicativo do Office deve suportar para que o código dentro da `if` instrução seja executado (por exemplo, "**1,9**").

> [!WARNING]
> Ao chamar o `isSetSupported` método, o valor do `MinimumVersion` parâmetro (se especificado) deverá ser uma cadeia de caracteres. Isso ocorre porque o analisador de JavaScript não pode diferenciar valores numéricos, como 1.1 e 1.10, onde é possível para valores de cadeia de caracteres como "1.1" e "1.10".
> A sobrecarga de `number` está obsoleta.

Use `isSetSupported` com o `RequirementSetName` associado com o aplicativo do Office da seguinte maneira.

|Aplicativo do Office|RequirementSetName|
|---|---|
|Excel|ExcelApi|
|OneNote|OneNoteApi|
|Outlook|Caixa de correio|
|Word|WordApi|

O `isSetSupported` método e os conjuntos de requisitos para esses aplicativos estão disponíveis no arquivo de Office.js mais recente na CDN. Se você não usar Office.js da CDN, seu suplemento poderá gerar exceções, pois `isSetSupported` será indefinido. Para obter mais informações, consulte [especificar a biblioteca de API JavaScript do Office mais recente](#specify-the-latest-office-javascript-api-library).

O exemplo de código a seguir mostra como um suplemento pode fornecer funcionalidades diferentes para diferentes aplicativos do Office que podem oferecer suporte a conjuntos de requisitos ou membros de API diferentes.

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

Alguns membros de API não pertencem a conjuntos de requisitos. Isso aplica-se somente a membros da API que fazem parte do namespace da [API JavaScript do Office](../reference/javascript-api-for-office.md) (qualquer coisa em relação à `Office.` exceção de [APIs de caixa de correio do Outlook](/javascript/api/outlook)), mas não membros da API que pertencem à [API JavaScript do Word](../reference/overview/word-add-ins-reference-overview.md) (qualquer coisa em `Word.` ), [API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md) (qualquer coisa em `Excel.` ) ou namespaces da [API JavaScript do OneNote](../reference/overview/onenote-add-ins-javascript-reference.md) (tudo em) `OneNote.` . Quando o suplemento depender de um método que não faz parte de um conjunto de requisitos, você poderá usar a verificação de tempo de execução para determinar se o método é compatível com o aplicativo do Office, conforme mostrado no exemplo de código a seguir. Para obter uma lista completa dos métodos que não pertencem a um conjunto de requisitos, confira [Conjuntos de requisitos de Suplemento do Office](../reference/requirement-sets/office-add-in-requirement-sets.md#methods-that-arent-part-of-a-requirement-set).

> [!NOTE]
> Recomendamos limitar o uso desse tipo de verificação no tempo de execução no código de seu suplemento.

O exemplo de código a seguir verifica se o aplicativo do Office suporta `document.setSelectedDataAsync` .

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
