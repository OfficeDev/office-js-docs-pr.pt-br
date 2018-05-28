---
title: Especificar hosts do Office e requisitos de API
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: bd517dee1faf8d3f3009a0b9ce7127f5760e730d
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="specify-office-hosts-and-api-requirements"></a>Especificar hosts do Office e requisitos de API

Seu Suplemento do Office pode depender de um host espec?fico do Office, um conjunto de requisitos, um membro de API ou uma vers?o da API para funcionar conforme o esperado. Por exemplo, o suplemento pode:

- Executar em um ?nico aplicativo do Office (Word ou Excel) ou diversos aplicativos.
    
- Usar as APIs de JavaScript que est?o dispon?veis apenas em algumas vers?es do Office. Por exemplo, voc? pode usar as APIs JavaScript do Excel em um suplemento executado no Excel 2016. 
    
- Executar apenas nas vers?es do Office que oferecem suporte a membros da API que seu suplemento usa.
    
Este artigo ajuda voc? a entender quais op??es voc? deve escolher para garantir que seu suplemento funcione conforme o esperado e atinja o p?blico mais amplo poss?vel.

> [!NOTE]
> Confira uma vis?o avan?ada da compatibilidade atual dos suplementos do Office no momento na p?gina [Disponibilidade de hosts e plataformas de suplementos do Office](../overview/office-add-in-availability.md). 

A tabela a seguir lista os principais conceitos discutidos neste artigo.

|**Conceito**|**Descri??o**|
|:-----|:-----|
|Aplicativo do Office, aplicativo host do Office, host do Office ou host|O aplicativo do Office usado para executar seu suplemento. Por exemplo, Word, Word Online, Excel etc.|
|Plataforma|Onde o host do Office ? executado, por exemplo, no Office Online ou no Office para iPad.|
|Conjunto de requisitos|Um grupo nomeado de membros relacionados da API. Os suplementos usam conjuntos de requisitos para determinar se o host do Office oferece suporte a membros da API usados por seu suplemento. ? mais f?cil testar se h? suporte para um conjunto de requisitos do que o suporte para membros individuais da API. O suporte a um conjunto de requisitos varia de acordo com o host do Office e a vers?o do host do Office. <br >Conjuntos de requisitos s?o especificados no arquivo de manifesto. Ao especificar conjuntos de requisitos no manifesto, voc? estabelece o n?vel m?nimo de suporte ? API que o host do Office deve fornecer a fim de executar seu suplemento. Os hosts do Office que n?o d?o suporte aos conjuntos de requisitos especificados no manifesto n?o podem executar o suplemento, e seu suplemento n?o aparecer? em <span class="ui">Meus Suplementos</span>. Isso restringe onde seu suplemento ? disponibilizado. Isso ? definido no c?digo usando verifica??es no tempo de execu??o. Para obter uma lista completa de conjuntos de requisitos, confira [Conjuntos de requisitos de Suplemento do Office](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).|
|Verifica??o no tempo de execu??o|Um teste ? executado no tempo de execu??o para determinar se o host do Office que est? executando seu suplemento oferece suporte aos conjuntos de requisitos ou m?todos usados por seu suplemento. Para executar uma verifica??o no tempo de execu??o, use uma instru??o **if** com o m?todo **isSetSupported**, os conjuntos de requisito ou os nomes de m?todo que n?o fazem parte de um conjunto de requisitos. Use as verifica??es no tempo de execu??o para garantir que seu suplemento alcance o maior n?mero de clientes. Ao contr?rio dos conjuntos de requisitos, as verifica??es no tempo de execu??o n?o especificam o n?vel m?nimo de suporte ? API exigido do host do Office para que seu suplemento possa ser executado. Em vez disso, use a instru??o **if** para determinar se h? suporte para um membro da API. Se houver, voc? poder? proporcionar mais funcionalidade em seu suplemento. Seu suplemento sempre aparecer? em **Meus Suplementos** ao usar verifica??es no tempo de execu??o.|

## <a name="before-you-begin"></a>Antes de come?ar

O suplemento deve usar a vers?o mais recente do esquema de manifesto de suplemento. Se voc? usar as verifica??es no tempo de execu??o em seu suplemento, use a biblioteca mais recente da API JavaScript para Office (office.js).

### <a name="specify-the-latest-add-in-manifest-schema"></a>Especificar o esquema de manifesto de suplemento mais recente

Seu manifesto de suplemento deve usar a vers?o 1.1 do esquema de manifesto de suplemento. Defina o elemento **OfficeApp** no manifesto do seu suplemento da seguinte maneira.

```XML
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
```

### <a name="specify-the-latest-javascript-api-for-office-library"></a>Especificar a biblioteca de API JavaScript para Office mais recente

Se voc? usar as verifica??es no tempo de execu??o, fa?a refer?ncia ? vers?o mais recente da biblioteca de API JavaScript para Office na CDN (rede de distribui??o de conte?do). Para tanto, adicione a seguinte marca `script` ao c?digo HTML. O uso de `/1/` na URL da CDN garante a refer?ncia ? vers?o mais recente do Office.js.

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

## <a name="options-to-specify-office-hosts-or-api-requirements"></a>Op??es para especificar os hosts do Office ou requisitos de API

Ao especificar os hosts do Office ou os requisitos de API, h? v?rios fatores a considerar. O diagrama a seguir mostra como decidir sobre qual t?cnica usar em seu suplemento.

![Escolha a melhor op??o para o seu suplemento ao especificar os hosts do Office ou os requisitos de API](../images/options-for-office-hosts.png)

- Se o seu suplemento for executado em um host do Office, defina o elemento **Hosts** no manifesto. Para saber mais, confira [Definir o elemento Hosts](#set-the-hosts-element).
    
- Para definir o conjunto de requisitos m?nimos ou os membros da API que devem receber suporte de um host do Office para que seu suplemento seja executado, defina o elemento **Requirements** no manifesto. Para saber mais, confira [Definir o elemento Requirements no manifesto](#set-the-requirements-element-in-the-manifest).
    
- Se voc? quiser fornecer outras funcionalidades caso conjuntos de requisitos ou membros da API espec?ficos estejam dispon?veis no host do Office, execute uma verifica??o no tempo de execu??o no c?digo JavaScript do seu suplemento. Por exemplo, se o seu suplemento for executado no Excel 2016, use os membros da nova API JavaScript para Excel a fim de fornecer outras funcionalidades. Para saber mais, confira [Usar verifica??es de tempo de execu??o em seu c?digo JavaScript](#use-runtime-checks-in-your-javascript-code).
    
## <a name="set-the-hosts-element"></a>Definir o elemento Hosts

Para fazer seu suplemento ser executado em um aplicativo host do Office, use os elementos  **Hosts** e **Host** no manifesto. Se voc? n?o especificar o elemento **Hosts**, o suplemento ser? executado em todos os hosts.

Por exemplo, a declara??o de **Hosts** e **Host** a seguir especifica que o suplemento funcionar? com qualquer vers?o do Excel, o que inclui o Excel para Windows, o Excel Online e o Excel para iPad.

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

O elemento **Hosts** pode conter um ou mais elementos **Host**. O elemento **Host** especifica o host do Office exigido por seu suplemento. O atributo **Name** ? obrigat?rio e pode ser definido com um dos valores a seguir.

| Nome          | Aplicativos host do Office                      |
|:--------------|:----------------------------------------------|
| Banco de dados      | Aplicativos Web do Access                               |
| Documento      | Word para Windows, Mac, iPad e Online        |
| Caixa de correio       | Outlook para Windows, Mac, Web e Outlook.com | 
| Apresenta??o  | PowerPoint para Windows, Mac, iPad e Online  |
| Projeto       | Projeto                                       |
| Pasta de trabalho      | Excel para Windows, Mac, iPad e Online           |

> [!NOTE]
> O atributo `Name` especifica o aplicativo host do Office que pode executar seu suplemento. H? suporte para hosts do Office em v?rias plataformas, que s?o executados em computadores, navegadores da Web, tablets e dispositivos m?veis. Voc? n?o pode especificar qual plataforma pode ser usada para executar seu suplemento. Por exemplo, se voc? especificar `Mailbox`, o Outlook e o Outlook Web App podem ser usados para executar o suplemento. 


## <a name="set-the-requirements-element-in-the-manifest"></a>Definir o elemento Requirements no manifesto

O elemento **Requirements** especifica os conjuntos de requisitos m?nimos ou os membros da API que devem receber suporte de um host do Office para que seu suplemento seja executado. O elemento **Requirements** pode especificar conjuntos de requisitos e m?todos individuais usados em seu suplemento. Na vers?o 1.1 do esquema de manifesto de suplemento, o elemento **Requirements** ? opcional para todos os suplementos, exceto para os suplementos do Outlook.

> [!WARNING]
> Use o elemento **Requirements** apenas para especificar conjuntos de requisitos ou membros de API cruciais ao seu suplemento. Se o host do Office ou a plataforma n?o der suporte ao conjunto de requisitos ou membros da API especificados no elemento **Requirements**, o suplemento n?o ser? executado no host ou na plataforma e n?o ser? exibido em **Meus Suplementos**. Em vez disso, recomendamos que voc? disponibilize seu suplemento em todas as plataformas de um host do Office, como o Excel para Windows, o Excel Online e o Excel para iPad. Para disponibilizar seu suplemento em _todos_ os hosts e plataformas do Office, use verifica??es no tempo de execu??o em vez do elemento **Requirements**.

O exemplo de c?digo a seguir mostra um suplemento que carrega em todos os aplicativos host do Office que oferecem suporte ao seguinte:

-  O conjunto de requisitos **TableBindings**, que tem uma vers?o m?nima de 1.1.
    
-  O conjunto de requisitos **OOXML**, que tem uma vers?o m?nima de 1.1.
    
-  O m?todo **Document.getSelectedDataAsync**.

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

- O elemento **Requirements** cont?m os elementos filhos **Sets** e **Methods**.
    
- O elemento **Sets** pode conter um ou mais elementos **Set**. **DefaultMinVersion** especifica o valor padr?o de **MinVersion** para todos os elementos filhos de **Set**.
    
- O elemento **Set** especifica os conjuntos de requisitos que devem receber suporte do host do Office para que o suplemento seja executado. O atributo **Name** especifica o nome do conjunto de requisitos. **MinVersion** especifica a vers?o m?nima do conjunto de requisitos. **MinVersion** substitui o valor de **DefaultMinVersion**. Para saber mais sobre os conjuntos de requisito e sobre as vers?es de conjuntos de requisitos aos quais membros de sua API pertencem, confira [Conjuntos de requisitos de suplementos do Office](https://dev.office.com/reference/add-ins/office-add-in-requirement-sets).
    
- O elemento **Methods** pode conter um ou mais elementos **Method**. Voc? n?o pode usar o elemento **Methods** com suplementos do Outlook.
    
- O elemento **Methods** especifica um m?todo individual que deve receber suporte no host do Office em que o suplemento ? executado. O atributo **Name** ? obrigat?rio e especifica o nome do m?todo qualificado com seu objeto pai.
    

## <a name="use-runtime-checks-in-your-javascript-code"></a>Usar verifica??es no tempo de execu??o em seu c?digo JavaScript


Se certos conjuntos de requisitos recebem suporte do host do Office, voc? pode proporcionar outras funcionalidades em seu suplemento. Por exemplo, pode usar a nova API JavaScript para Word em seu suplemento existente se o seu suplemento for executado no Word 2016. Para fazer isso, use o m?todo **isSetSupported** com o nome do conjunto de requisitos. **isSetSupported** determinado, no tempo de execu??o, se o host do Office que est? executando o suplemento d? suporte ao conjunto de requisitos. Se houver suporte para o conjunto de requisitos, **isSetSupported** retorna **true** e executa o c?digo adicional que usa os membros da API desse conjunto de requisitos. Se o host do Office n?o d? suporte ao conjunto de requisitos, **isSetSupported** retorna **false** e o c?digo adicional n?o ? executado. O c?digo a seguir mostra a sintaxe a ser usada com **isSetSupported**.


```js
if (Office.context.requirements.isSetSupported(RequirementSetName , VersionNumber))
{
   // Code that uses API members from RequirementSetName.
}

```


-  _RequirementSetName_ (obrigat?rio) ? uma cadeia de caracteres que representa o nome do conjunto de requisitos. Para saber mais sobre os conjuntos de requisitos dispon?veis, confira [Conjuntos de requisitos de Suplemento do Office](https://dev.office.com/reference/add-ins/office-add-in-requirement-sets).
    
-  _VersionNumber_ (opcional) ? a vers?o do conjunto de requisitos.
    
No Excel 2016 ou no Word 2016, use **isSetSupported** com os conjuntos de requisitos **ExcelAPI** ou **WordAPI**. O m?todo **isSetSupported** e os conjuntos de requisitos **ExcelAP**I e **WordAPI** est?o dispon?veis no Office.js mais recente na CDN. Se voc? n?o usa o Office.js da CDN, seu suplemento pode gerar exce??es, pois **isSetSupported** fica indefinido. Para saber mais, confira [Especificar a biblioteca de API JavaScript para Office mais recente](#specify-the-latest-javascript-api-for-office-library). 


> [!NOTE]
> **isSetSupported** n?o funciona no Outlook ou no Outlook Web App. Para usar uma verifica??o no tempo de execu??o no Outlook ou no Outlook Web App, use a t?cnica descrita em [Verifica??es no tempo de execu??o usando m?todos que n?o fazem parte de um conjunto de requisitos](#runtime-checks-using-methods-not-in-a-requirement-set).

O exemplo de c?digo a seguir mostra como um suplemento pode fornecer outras funcionalidades para hosts do Office diferentes que podem dar suporte a conjuntos de requisitos ou membros de API diferentes.




```js
if (Office.context.requirements.isSetSupported('WordApi', 1.1))
{
    // Run code that provides additional functionality using the JavaScript API for Word when the add-in runs in Word 2016.
}
else if (Office.context.requirements.isSetSupported('CustomXmlParts'))
{
      // Run code that uses API members from the CustomXmlParts requirement set.
}
else 
{
    // Run additional code when the Office host is not Word 2016, and when the Office host does not support the CustomXmlParts requirement set.
}

```


## <a name="runtime-checks-using-methods-not-in-a-requirement-set"></a>Verifica??es no tempo de execu??o usando m?todos que n?o fazem parte de um conjunto de requisitos


Alguns membros de API n?o pertencem a conjuntos de requisitos. Isso aplica-se apenas a membros da API que fazem parte do namespace [API JavaScript para Office](https://dev.office.com/reference/add-ins/javascript-api-for-office) (qualquer coisa abaixo de Office.), n?o a membros de API que pertencem a namespaces da API JavaScript para Word (qualquer coisa em Word.) ou da [Refer?ncia sobre a API JavaScript para suplementos do Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview) (qualquer coisa em Excel.). Quando seu suplemento depende de um m?todo que n?o faz parte de um conjunto de requisitos, ? poss?vel usar a verifica??o no tempo de execu??o para determinar se o m?todo tem suporte no host do Office, conforme mostra o exemplo de c?digo a seguir. Para obter uma lista completa dos m?todos que n?o pertencem a um conjunto de requisitos, confira [Conjuntos de requisitos de Suplemento do Office](https://dev.office.com/reference/add-ins/office-add-in-requirement-sets).


> [!NOTE]
> Recomendamos limitar o uso desse tipo de verifica??o no tempo de execu??o no c?digo de seu suplemento.

O exemplo de c?digo a seguir verifica se o host oferece suporte a **document.setSelectedDataAsync**.




```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses document.setSelectedDataAsync.
}
```


## <a name="see-also"></a>Veja tamb?m

- [Manifesto XML dos Suplementos do Office](add-in-manifests.md)
- [Conjuntos de requisitos de Suplemento do Office](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)
- [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)