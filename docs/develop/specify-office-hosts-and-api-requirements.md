---
title: Especificar hosts do Office e requisitos de API
description: Saiba como especificar os requisitos de API e aplicativos do Office para que seu suplemento funcione conforme o esperado.
ms.date: 05/19/2022
ms.localizationpriority: medium
ms.openlocfilehash: 60d69c9fae136e73bf9920c7dc96f13d832331fd
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810292"
---
# <a name="specify-office-applications-and-api-requirements"></a>Especificar requisitos da API e de aplicativos do Office

O Suplemento do Office pode depender de um aplicativo específico do Office (também chamado de host do Office) ou de membros específicos da API JavaScript do Office (office.js). Por exemplo, o suplemento pode:

- Executar em um único aplicativo do Office (por exemplo, Word ou Excel) ou diversos aplicativos.
- Use as APIs JavaScript do Office que só estão disponíveis em algumas versões do Office. Por exemplo, a versão perpétua licenciada por volume do Excel 2016 não dá suporte a todas as APIs relacionadas ao Excel na biblioteca JavaScript do Office.

Nessas situações, você precisa garantir que seu suplemento nunca seja instalado em aplicativos do Office ou versões do Office nas quais ele não possa ser executado.

Há também cenários em que você deseja controlar quais recursos do suplemento estão visíveis para os usuários com base em seu aplicativo do Office e na versão do Office. Dois exemplos são:

- Seu suplemento tem recursos úteis no Word e no PowerPoint, como manipulação de texto, mas tem alguns recursos adicionais que só fazem sentido no PowerPoint, como recursos de gerenciamento de slides. Você precisa ocultar os recursos somente do PowerPoint quando o suplemento estiver em execução no Word.
- Seu suplemento tem um recurso que requer um método de API JavaScript do Office com suporte em algumas versões de um aplicativo do Office, como o Excel de assinatura do Microsoft 365, mas não tem suporte em outras, como Excel 2016 perpétuos licenciados por volume. Mas seu suplemento tem outros recursos que exigem apenas métodos de API JavaScript do *Office com suporte* em Excel 2016 perpétuos licenciados por volume. Nesse cenário, você precisa que o suplemento seja instalado nessa versão do Excel 2016, mas o recurso que exige o método sem suporte deve ser oculto desses usuários.

Este artigo ajuda você a entender quais opções você deve escolher para garantir que seu suplemento funcione conforme o esperado e atinja o público mais amplo possível.

> [!NOTE]
> Para obter uma exibição de alto nível de onde os Suplementos do Office têm suporte no momento, consulte a página [aplicativo cliente do Office e a disponibilidade da plataforma para suplementos do Office](/javascript/api/requirement-sets) .

> [!TIP]
> Muitas das tarefas descritas neste artigo são feitas para você, no todo ou em parte, quando você cria seu projeto de suplemento com uma ferramenta, como o [gerador Yeoman para Suplementos do Office](yeoman-generator-overview.md) ou um dos modelos de Suplemento do Office no Visual Studio. Nesses casos, interprete a tarefa como o significado de que você deve verificar se ela foi feita.

## <a name="use-the-latest-office-javascript-api-library"></a>Usar a biblioteca de API JavaScript mais recente

Seu suplemento deve carregar a versão mais atual da biblioteca de API JavaScript do Office da CDN (rede de entrega de conteúdo). Para fazer isso, certifique-se de ter a marca a seguir `script` no primeiro arquivo HTML que seu suplemento abre. O uso de `/1/` na URL da CDN garante a referência à versão mais recente do Office.js.

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## <a name="specify-which-office-applications-can-host-your-add-in"></a>Especificar quais aplicativos do Office podem hospedar seu suplemento

Por padrão, um suplemento é instalável em todos os aplicativos do Office com suporte pelo tipo de suplemento especificado (ou seja, Email, painel de tarefas ou Conteúdo). Por exemplo, um suplemento de painel de tarefas pode ser instalado por padrão no Access, Excel, OneNote, PowerPoint, Project e Word.

Para garantir que o suplemento seja instalável em um subconjunto de aplicativos do Office, use os elementos [Hosts](/javascript/api/manifest/hosts) e [Host](/javascript/api/manifest/host) no manifesto.

Por exemplo, o seguinte **\<Hosts\>** e **\<Host\>** a declaração especificam que o suplemento pode ser instalado em qualquer versão do Excel, que inclui Excel na Web, Windows e iPad, mas não pode ser instalado em nenhum outro aplicativo do Office.

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

O **\<Hosts\>** elemento pode conter um ou mais **\<Host\>** elementos. Deve haver um elemento separado **\<Host\>** para cada aplicativo do Office no qual o suplemento deve ser instalável. O `Name` atributo é necessário e pode ser definido como um dos valores a seguir.

| Nome          | Aplicativos cliente do Office                     | Tipos de suplemento disponíveis |
|:--------------|:-----------------------------------------------|:-----------------------|
| Banco de dados      | Aplicativos Web do Access                                | Painel de tarefas              |
| Documento      | Word na Web, Windows, Mac, iPad            | Painel de tarefas              |
| Mailbox       | Outlook na Web, Windows, Mac, Android, iOS | Email                   |
| Notebook      | OneNote Online                             | Painel de tarefas, Conteúdo     |
| Presentation  | PowerPoint na Web, Windows, Mac, iPad      | Painel de tarefas, Conteúdo     |
| Project       | Project no Windows                             | Painel de tarefas              |
| Pasta de Trabalho      | Excel na Web, Windows, Mac, iPad           | Painel de tarefas, Conteúdo     |

> [!NOTE]
> Os aplicativos do Office têm suporte em diferentes plataformas e são executados em desktops, navegadores da Web, tablets e dispositivos móveis. Normalmente, você não pode especificar qual plataforma pode ser usada para executar seu suplemento. Por exemplo, se você especificar `Workbook`, tanto Excel na Web quanto no Windows podem ser usados para executar seu suplemento. No entanto, se você especificar `Mailbox`, seu suplemento não será executado em clientes móveis do Outlook, a menos que você defina o [ponto de extensão móvel](/javascript/api/manifest/extensionpoint#mobilemessagereadcommandsurface).

> [!NOTE]
> Não é possível que um manifesto de suplemento se aplique a mais de um tipo: Email, painel de tarefas ou Conteúdo. Isso significa que, se você quiser que seu suplemento seja instalado no Outlook e em um dos outros aplicativos do Office, você deve criar *dois* suplementos, um com um manifesto de tipo mail e outro com um painel de tarefas ou manifesto de tipo de conteúdo.

> [!IMPORTANT]
> Não recomendamos mais criar e usar aplicativos Web do Access e bancos de dados no SharePoint. Como alternativa, use o [Microsoft PowerApps](https://powerapps.microsoft.com/) para criar soluções de negócios sem código para dispositivos móveis e Web.

## <a name="specify-which-office-versions-and-platforms-can-host-your-add-in"></a>Especifique quais versões e plataformas do Office podem hospedar seu suplemento

Você não pode especificar explicitamente as versões e builds do Office ou as plataformas nas quais seu suplemento deve ser instalável, e você não gostaria de, porque você teria que revisar seu manifesto sempre que o suporte para os recursos de suplemento que seu suplemento usa for estendido para uma nova versão ou plataforma. Em vez disso, especifique no manifesto as APIs que seu suplemento precisa. O Office impede que o suplemento seja instalado em combinações de versão e plataforma do Office que não dão suporte às APIs e garante que o suplemento não apareça **em Meus Suplementos**.

> [!IMPORTANT]
> Use apenas o manifesto base para especificar os membros da API que seu suplemento deve ter que ser de qualquer valor significativo. Se o suplemento usa uma API para alguns recursos, mas tem outros recursos úteis que não exigem a API, você deve projetar o suplemento para que ele seja instalado em combinações de versão da plataforma e do Office que não dão suporte à API, mas fornece uma experiência reduzida nessas combinações. Para obter mais informações, consulte [Design para experiências alternativas](#design-for-alternate-experiences).

### <a name="requirement-sets"></a>Conjuntos de requisitos

Para simplificar o processo de especificação das APIs que seu suplemento precisa, o Office agrupa a maioria das APIs em *conjuntos de requisitos*. As APIs no [Modelo de Objeto de API Comum](understanding-the-javascript-api-for-office.md#api-models) são agrupadas pelo recurso de desenvolvimento que elas dão suporte. Por exemplo, todas as APIs conectadas às associações de tabela estão no conjunto de requisitos chamado "TableBindings 1.1". As APIs nos [modelos de objeto específicos do aplicativo](understanding-the-javascript-api-for-office.md#api-models) são agrupadas quando são lançadas para uso em suplementos de produção.

Os conjuntos de requisitos são versão. Por exemplo, as APIs que dão suporte a [Caixas de Diálogo](../develop/dialog-api-in-office-add-ins.md) estão no conjunto de requisitos DialogApi 1.1. Quando APIs adicionais que habilitam mensagens de um painel de tarefas para uma caixa de diálogo foram lançadas, elas foram agrupadas em DialogApi 1.2, juntamente com todas as APIs no DialogApi 1.1. *Cada versão de um conjunto de requisitos é um superconjunto de todas as versões anteriores.*

O suporte ao conjunto de requisitos varia de acordo com o aplicativo do Office, a versão do aplicativo do Office e a plataforma na qual ele está em execução. Por exemplo, o DialogApi 1.2 não tem suporte em versões perpétuas licenciadas por volume do Office antes do Office 2021, mas o DialogApi 1.1 tem suporte em todas as versões perpétuas de volta ao Office 2013. Você deseja que seu suplemento seja instalado em todas as combinações de plataforma e versão do Office que dão suporte às APIs que ele usa, portanto, você deve sempre especificar no manifesto a versão *mínima* de cada conjunto de requisitos exigido pelo suplemento. Os detalhes sobre como fazer isso são posteriores neste artigo.

> [!TIP]
> Para obter mais informações sobre a versão do conjunto de requisitos, consulte [Disponibilidade dos conjuntos de requisitos do Office](office-versions-and-requirement-sets.md#office-requirement-sets-availability) e para as listas completas de conjuntos de requisitos e informações sobre as APIs em cada, comece com [conjuntos de requisitos do Suplemento do Office](/javascript/api/requirement-sets/common/office-add-in-requirement-sets). Os tópicos de referência para a maioria das APIs Office.js também especificam o conjunto de requisitos ao qual pertencem (se houver).

> [!NOTE]
> Alguns conjuntos de requisitos também têm elementos de manifesto associados a eles. Consulte [Especificando requisitos em um elemento VersionOverrides](#specify-requirements-in-a-versionoverrides-element) para obter informações sobre quando esse fato é relevante para o design do suplemento.

#### <a name="apis-not-in-a-requirement-set"></a>APIs não em um conjunto de requisitos

Todas as APIs nos modelos específicos do aplicativo estão em conjuntos de requisitos, mas algumas delas no modelo de API Comum não estão. Há também uma maneira de especificar uma dessas APIs sem conjunto no manifesto quando o suplemento exigir uma. Os detalhes estão incluídos mais adiante neste artigo.

### <a name="requirements-element"></a>Elemento Requirements

Use o elemento [Requisitos](/javascript/api/manifest/requirements) e seus elementos filho [Conjuntos](/javascript/api/manifest/sets) e [Métodos](/javascript/api/manifest/methods) para especificar os conjuntos de requisitos mínimos ou membros da API que devem ter suporte pelo aplicativo do Office para instalar seu suplemento.

Se o aplicativo ou plataforma do Office não der suporte aos conjuntos de requisitos ou membros da API especificados no **\<Requirements\>** elemento, o suplemento não será executado nesse aplicativo ou plataforma e não será exibido **em Meus Suplementos**.

> [!NOTE]
> O **\<Requirements\>** elemento é opcional para todos os suplementos, exceto para suplementos do Outlook. Quando o `xsi:type` atributo do elemento raiz `OfficeApp` é `MailApp`, deve haver um **\<Requirements\>** elemento que especifica a versão mínima do conjunto de requisitos da caixa de correio que o suplemento requer. Para obter mais informações, consulte [Conjuntos de requisitos da API JavaScript do Outlook](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets).

O exemplo de código a seguir mostra como configurar um suplemento que é instalável em todos os aplicativos do Office que dão suporte ao seguinte:

- `TableBindings` conjunto de requisitos, que tem uma versão mínima de "1.1".
- `OOXML` conjunto de requisitos, que tem uma versão mínima de "1.1".
- `Document.getSelectedDataAsync` Método.

```XML
<OfficeApp ... >
  ...
  <Requirements>
     <Sets DefaultMinVersion="1.1">
        <Set Name="TableBindings" MinVersion="1.1"/>
        <Set Name="OOXML" MinVersion="1.1"/>
     </Sets>
     <Methods>
        <Method Name="Document.getSelectedDataAsync"/>
     </Methods>
  </Requirements>
    ...
</OfficeApp>
```

Observe o seguinte sobre este exemplo.

- O **\<Requirements\>** elemento contém os **\<Sets\>** elementos filho e **\<Methods\>** .
- O **\<Sets\>** elemento pode conter um ou mais **\<Set\>** elementos. `DefaultMinVersion` especifica o valor padrão `MinVersion` de todos os elementos filho **\<Set\>** .
- Um elemento [Set](/javascript/api/manifest/set) especifica um conjunto de requisitos que o aplicativo do Office deve dar suporte para tornar o suplemento instalável. O `Name` atributo especifica o nome do conjunto de requisitos. A `MinVersion` especifica a versão mínima do conjunto de requisitos. `MinVersion` substitui o valor do `DefaultMinVersion` atributo no pai **\<Sets\>**.
- O **\<Methods\>** elemento pode conter um ou mais elementos [de método](/javascript/api/manifest/method) . Você não pode usar o **\<Methods\>** elemento com suplementos do Outlook.
- O **\<Method\>** elemento especifica um método individual que o aplicativo do Office deve dar suporte para tornar o suplemento instalável. O `Name` atributo é necessário e especifica o nome do método qualificado com seu objeto pai.

## <a name="design-for-alternate-experiences"></a>Design para experiências alternativas

Os recursos de extensibilidade que a plataforma suplemento do Office fornece podem ser divididos de maneira útil em três tipos:

- Recursos de extensibilidade que estão disponíveis imediatamente após a instalação do suplemento. Você pode usar esse tipo de recurso configurando um elemento [VersionOverrides](/javascript/api/manifest/versionoverrides) no manifesto. Um exemplo desse tipo de recurso é [Comandos de Suplemento](../design/add-in-commands.md), que são botões e menus personalizados de faixa de opções.
- Recursos de extensibilidade disponíveis somente quando o suplemento está em execução e que são implementados com Office.js APIs JavaScript; por exemplo, [Caixas de diálogo](../develop/dialog-api-in-office-add-ins.md).
- Recursos de extensibilidade que estão disponíveis apenas no runtime, mas são implementados com uma combinação de Office.js JavaScript e configuração em um **\<VersionOverrides\>** elemento. Exemplos delas são [funções personalizadas do Excel](../excel/custom-functions-overview.md), [logon único](sso-in-office-add-ins.md) e [guias contextuais personalizadas](../design/contextual-tabs.md).

Se o suplemento usar um recurso de extensibilidade específico para algumas de suas funcionalidades, mas tiver outra funcionalidade útil que não exija o recurso de extensibilidade, você deverá projetar o suplemento para que ele seja instalado em combinações de versão da plataforma e do Office que não dão suporte ao recurso de extensibilidade. Ele pode fornecer uma experiência valiosa, embora diminuída, nessas combinações.

Você implementa esse design de forma diferente dependendo de como o recurso de extensibilidade é implementado:

- Para obter recursos implementados inteiramente com JavaScript, consulte [Verificações do Runtime para obter suporte ao método e ao conjunto de requisitos](#runtime-checks-for-method-and-requirement-set-support).
- Para recursos que exigem que você configure um **\<VersionOverrides\>** elemento, consulte [Especificando requisitos em um elemento VersionOverrides](#specify-requirements-in-a-versionoverrides-element).

### <a name="runtime-checks-for-method-and-requirement-set-support"></a>Verificações de runtime para suporte ao método e ao conjunto de requisitos

Você testa no runtime para descobrir se o Office do usuário dá suporte a um conjunto de requisitos com o método [isSetSupported](/javascript/api/office/office.requirementsetsupport#office-office-requirementsetsupport-issetsupported-member(1)) . Passe o nome do conjunto de requisitos e a versão mínima como parâmetros. Se o conjunto de requisitos tiver suporte, `isSetSupported` retornará `true`. O código a seguir mostra um exemplo.

```js
if (Office.context.requirements.isSetSupported('WordApi', '1.1'))
{
   // Code that uses API members from the WordApi 1.1 requirement set.
} else {
   // Provide diminished experience here. E.g., run alternate code when the user's Word is perpetual Word 2013 (which does not support WordApi 1.1).
}
```

Sobre este código, observe:

- O primeiro parâmetro é necessário. É uma cadeia de caracteres que representa o nome do conjunto de requisitos. Para saber mais sobre os conjuntos de requisitos disponíveis, confira [Conjuntos de requisitos de Suplemento do Office](/javascript/api/requirement-sets/common/office-add-in-requirement-sets).
- O segundo parâmetro é opcional. É uma cadeia de caracteres que especifica a versão mínima do conjunto de requisitos que o aplicativo do Office deve dar suporte para que o código dentro da `if` instrução seja executado (por exemplo, "**1.9**"). Se não for usada, a versão "1.1" será assumida.

> [!WARNING]
> Ao chamar o `isSetSupported` método, o valor do segundo parâmetro (se especificado) deve ser uma cadeia de caracteres e não um número. O analisador JavaScript não pode diferenciar entre valores numéricos como 1.1 e 1.10, enquanto pode para valores de cadeia de caracteres como "1.1" e "1.10".

A tabela a seguir mostra os nomes de conjunto de requisitos para os modelos de API específicos do aplicativo.

|Aplicativo do Office|RequirementSetName|
|---|---|
|Excel|ExcelApi|
|OneNote|OneNoteApi|
|Outlook|Caixa de correio|
|PowerPoint|PowerPointApi|
|Word|WordApi|

Veja a seguir um exemplo de como usar o método com um dos conjuntos de requisitos de modelo de API Comum.

```js
if (Office.context.requirements.isSetSupported('CustomXmlParts'))
{
    // Run code that uses API members from the CustomXmlParts requirement set.
}
else
{
    // Run alternate code when the user's Office application doesn't support the CustomXmlParts requirement set.
}
```

> [!NOTE]
> O `isSetSupported` método e os conjuntos de requisitos para esses aplicativos estão disponíveis no arquivo de Office.js mais recente na CDN. Se você não usar Office.js da CDN, seu suplemento poderá gerar exceções se você estiver usando uma versão antiga da biblioteca na qual `isSetSupported` está indefinida. Para obter mais informações, consulte [Usar a biblioteca de API JavaScript mais recente do Office](#use-the-latest-office-javascript-api-library).

Quando o suplemento depender de um método que não faz parte de um conjunto de requisitos, use a verificação de runtime para determinar se o método tem suporte pelo aplicativo do Office, conforme mostrado no exemplo de código a seguir. Para obter uma lista completa dos métodos que não pertencem a um conjunto de requisitos, confira [Conjuntos de requisitos de Suplemento do Office](/javascript/api/requirement-sets/common/office-add-in-requirement-sets#methods-that-arent-part-of-a-requirement-set).

> [!NOTE]
> Recomendamos limitar o uso desse tipo de verificação no tempo de execução no código de seu suplemento.

O exemplo de código a seguir verifica se o aplicativo do Office dá `document.setSelectedDataAsync`suporte a .

```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses `document.setSelectedDataAsync`.
}
```

### <a name="specify-requirements-in-a-versionoverrides-element"></a>Especificar requisitos em um elemento VersionOverrides

O elemento [VersionOverrides](/javascript/api/manifest/versionoverrides) foi adicionado ao esquema de manifesto principalmente, mas não exclusivamente, para dar suporte a recursos que devem estar disponíveis imediatamente após a instalação de um suplemento, como comandos de suplemento (botões e menus personalizados). O Office deve saber sobre esses recursos quando analisar o manifesto de suplemento.

Suponha que o suplemento use um desses recursos, mas o suplemento é valioso e deve ser instalável, mesmo em versões do Office que não dão suporte ao recurso. Nesse cenário, identifique o recurso usando um elemento [Requirements](/javascript/api/manifest/requirements) (e seus elementos [conjuntos](/javascript/api/manifest/sets) e [métodos](/javascript/api/manifest/methods) filho) que você inclui como filho do **\<VersionOverrides\>** elemento em si, em vez de como filho do elemento base `OfficeApp` . O efeito de fazer isso é que o Office permitirá que o suplemento seja instalado, mas o Office ignorará alguns dos elementos filho do **\<VersionOverrides\>** elemento nas versões do Office em que o recurso não tem suporte.

Especificamente, os elementos filho dos **\<VersionOverrides\>** elementos que substituem no manifesto base, como um **\<Hosts\>** elemento, são ignorados e os elementos correspondentes do manifesto base são usados. No entanto, pode haver elementos filho em um **\<VersionOverrides\>** que realmente implemente recursos adicionais em vez de substituir configurações no manifesto base. Dois exemplos são o `WebApplicationInfo` e `EquivalentAddins`. Essas partes do **\<VersionOverrides\>** *não* serão ignoradas, supondo que a plataforma e a versão do Office dão suporte ao recurso correspondente.  

Para obter informações sobre os elementos descendentes do **\<Requirements\>** elemento, consulte [Elemento Requisitos](#requirements-element) anteriormente neste artigo.

Apresentamos um exemplo a seguir.

```XML
<VersionOverrides ... >
   ...
   <Requirements>
      <Sets DefaultMinVersion="1.1">
         <Set Name="WordApi" MinVersion="1.2"/>
      </Sets>
   </Requirements>
   <Hosts>

      <!-- ALL MARKUP INSIDE THE HOSTS ELEMENT IS IGNORED WHEREVER WordApi 1.2 IS NOT SUPPORTED -->

      <Host xsi:type="Workbook">
         <!-- markup for custom add-in commands -->
      </Host>
   </Hosts>
</VersionOverrides>
```

> [!WARNING]
> Use muito cuidado antes de incluir um **\<Requirements\>** elemento em um **\<VersionOverrides\>**, porque em combinações de plataforma e versão que não dão suporte ao requisito, *nenhum* dos comandos de suplemento será instalado, *mesmo aqueles que invocam a funcionalidade que não precisa do requisito*. Considere, por exemplo, um suplemento que tenha dois botões de faixa de opções personalizados. Uma delas chama APIs JavaScript do Office que estão disponíveis no conjunto de requisitos **ExcelApi 1.4** (e posterior). As outras APIs de chamadas que só estão disponíveis no **ExcelApi 1.9** (e posterior). Se você colocar um requisito para o **\<VersionOverrides\>****ExcelApi 1.9** no , quando 1.9 não tiver suporte *, nenhum dos botões* aparecerá na faixa de opções. Uma estratégia melhor nesse cenário seria usar a técnica descrita em [verificações do Runtime para o método e o suporte ao conjunto de requisitos](#runtime-checks-for-method-and-requirement-set-support). O código invocado pelo segundo botão usa primeiro `isSetSupported` para verificar o suporte do **ExcelApi 1.9**. Se não houver suporte, o código fornecerá ao usuário uma mensagem dizendo que esse recurso do suplemento não está disponível na versão do Office.

> [!TIP]
> Não adianta repetir um elemento **Requirement** em um **\<VersionOverrides\>** que já aparece no manifesto base. Se o requisito for especificado no manifesto base, o suplemento não poderá instalar onde o requisito não tem suporte para que o Office nem analise o **\<VersionOverrides\>** elemento.

## <a name="see-also"></a>Confira também

- [Manifesto XML dos Suplementos do Office](add-in-manifests.md)
- [Conjuntos de requisitos de Suplemento do Office](/javascript/api/requirement-sets/common/office-add-in-requirement-sets)
- [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)
