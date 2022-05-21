---
title: Especificar hosts do Office e requisitos de API
description: Saiba como especificar Office aplicativos e requisitos de API para que o suplemento funcione conforme o esperado.
ms.date: 05/19/2022
ms.localizationpriority: medium
ms.openlocfilehash: 60ad00c918b04b6f12ecb6eec6c40772448b2ab8
ms.sourcegitcommit: 4ca3334f3cefa34e6b391eb92a429a308229fe89
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2022
ms.locfileid: "65628044"
---
# <a name="specify-office-applications-and-api-requirements"></a>Especificar requisitos da API e de aplicativos do Office

O Office do Office pode depender de um aplicativo Office específico (também chamado de host do Office) ou de membros específicos da API JavaScript (office.js) do Office. Por exemplo, o suplemento pode:

- Executar em um único aplicativo do Office (por exemplo, Word ou Excel) ou diversos aplicativos.
- Use as Office APIs JavaScript que só estão disponíveis em algumas versões do Office. Por exemplo, a versão de compra única do Excel 2016 não dá suporte Excel APIs relacionadas a Office biblioteca JavaScript.

Nessas situações, você precisa garantir que o suplemento nunca esteja instalado em aplicativos Office ou em versões Office em que ele não pode ser executado.

Também há cenários em que você deseja controlar quais recursos do seu suplemento são visíveis para os usuários com base em seu aplicativo Office e Office versão. Dois exemplos são:

- Seu suplemento tem recursos que são úteis no Word e no PowerPoint, como manipulação de texto, mas tem alguns recursos adicionais que só fazem sentido no PowerPoint, como recursos de gerenciamento de slides. Você precisa ocultar os PowerPoint somente quando o suplemento estiver em execução no Word.
- Seu suplemento tem um recurso que requer um método de API JavaScript do Office que tem suporte em algumas versões de um aplicativo do Office, como o Excel de assinatura, mas não tem suporte em outras pessoas, como compra única Excel 2016. Mas seu suplemento tem outros recursos que exigem apenas Office de API JavaScript com suporte no Excel 2016. Nesse cenário, você precisa que o suplemento seja instalável no Excel 2016, mas o recurso que requer o método sem suporte deve estar oculto dos usuários de Excel 2016.

Este artigo ajuda você a entender quais opções você deve escolher para garantir que seu suplemento funcione conforme o esperado e atinja o público mais amplo possível.

> [!NOTE]
> Para obter uma exibição de alto nível de onde Office suplementos têm suporte no momento, consulte o aplicativo cliente Office e a disponibilidade da plataforma para Office de [suplementos](/javascript/api/requirement-sets).

> [!TIP]
> Muitas das tarefas descritas neste artigo são feitas para você, no todo ou em parte, quando você cria seu projeto de suplemento com uma ferramenta, como o gerador [Yeoman para suplementos do Office](yeoman-generator-overview.md) ou um dos modelos de suplemento do Office no Visual Studio. Nesses casos, interprete a tarefa como o que significa que você deve verificar se ela foi feita.

## <a name="use-the-latest-office-javascript-api-library"></a>Usar a biblioteca Office API JavaScript mais recente

O suplemento deve carregar a versão mais atual da biblioteca Office API JavaScript da rede de distribuição de conteúdo (CDN). Para fazer isso, verifique se você tem a marca a `script` seguir no primeiro arquivo HTML que seu suplemento abre. O uso de `/1/` na URL da CDN garante a referência à versão mais recente do Office.js.

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## <a name="specify-which-office-applications-can-host-your-add-in"></a>Especificar quais Office aplicativos podem hospedar seu suplemento

Por padrão, um suplemento pode ser instalado em todos os Office aplicativos compatíveis com o tipo de suplemento especificado (ou seja, Email, Painel de Tarefas ou Conteúdo). Por exemplo, um suplemento do painel de tarefas pode ser instalado por padrão no Access, Excel, OneNote, PowerPoint, Project e Word. 

Para garantir que o suplemento seja inserível em um subconjunto de Office aplicativos, use os elementos [Hosts](/javascript/api/manifest/hosts) e [Host](/javascript/api/manifest/host) no manifesto.

Por exemplo, a declaração **Hosts** e **Host** a seguir especifica que o suplemento pode ser instalado em qualquer versão do Excel, que inclui Excel na Web, Windows e iPad, mas não pode ser instalado em nenhum outro aplicativo Office.

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

O **elemento Hosts** pode conter um ou mais **elementos host** . Deve haver um elemento **Host** separado para cada Office aplicativo no qual o suplemento deve ser instalado. O `Name` atributo é necessário e pode ser definido como um dos valores a seguir.

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
> Office aplicativos têm suporte em diferentes plataformas e são executados em desktops, navegadores da Web, tablets e dispositivos móveis. Normalmente, você não pode especificar qual plataforma pode ser usada para executar o suplemento. Por exemplo, se você especificar `Workbook`, Excel na Web e em Windows podem ser usados para executar o suplemento. No entanto, se você especificar`Mailbox`, o suplemento não será executado em Outlook clientes móveis, a menos que você defina o [ponto de extensão móvel](/javascript/api/manifest/extensionpoint#mobilemessagereadcommandsurface).

> [!NOTE]
> Não é possível que um manifesto de suplemento se aplique a mais de um tipo: Email, Painel de tarefas ou Conteúdo. Isso significa que, se você quiser que o suplemento seja instalado no Outlook e em um dos outros aplicativos do Office, deverá criar dois suplementos, um com um  manifesto de tipo Email e outro com um painel tarefa ou manifesto do tipo Conteúdo.

> [!IMPORTANT]
> Não recomendamos mais criar e usar aplicativos Web do Access e bancos de dados no SharePoint. Como alternativa, use o [Microsoft PowerApps](https://powerapps.microsoft.com/) para criar soluções de negócios sem código para dispositivos móveis e Web.

## <a name="specify-which-office-versions-and-platforms-can-host-your-add-in"></a>Especificar quais Office e plataformas podem hospedar seu suplemento

Você não pode especificar explicitamente as versões e builds do Office ou as plataformas nas quais o suplemento deve ser instalado, e você não gostaria de fazer isso porque precisaria revisar seu manifesto sempre que o suporte para os recursos de suplemento usados pelo suplemento fosse estendido para uma nova versão ou plataforma. Em vez disso, especifique no manifesto as APIs de que seu suplemento precisa. Office impede que o suplemento seja instalado em combinações de uma versão e plataforma do Office que não dão suporte às APIs e garante que o suplemento não aparecerá em Meus **Suplementos**.

> [!IMPORTANT]
> Use apenas o manifesto base para especificar os membros da API que seu suplemento deve ter de qualquer valor significativo. Se o suplemento usa uma API para alguns recursos, mas tem outros recursos úteis que não exigem a API, você deve projetar o suplemento para que ele possa ser instalado na plataforma e combinações de versão do Office que não dão suporte à API, mas fornece uma experiência diminuída nessas combinações. Para obter mais informações, consulte [Design para experiências alternativas](#design-for-alternate-experiences).

### <a name="requirement-sets"></a>Conjuntos de requisitos

Para simplificar o processo de especificação das APIs de que seu suplemento precisa, Office agrupa a maioria das APIs em conjuntos *de requisitos*. As APIs no [Modelo de Objeto de API](understanding-the-javascript-api-for-office.md#api-models) Comum são agrupadas pelo recurso de desenvolvimento ao qual elas dão suporte. Por exemplo, todas as APIs conectadas a associações de tabela estão no conjunto de requisitos chamado "TableBindings 1.1". As APIs nos modelos [de](understanding-the-javascript-api-for-office.md#api-models) objeto específicos do aplicativo são agrupadas por quando elas foram liberadas para uso em suplementos de produção.

Os conjuntos de requisitos têm controle de versão. Por exemplo, as APIs que dão suporte [a Caixas de](../develop/dialog-api-in-office-add-ins.md) Diálogo estão no conjunto de requisitos DialogApi 1.1. Quando apIs adicionais que habilitam mensagens de um painel de tarefas para uma caixa de diálogo foram lançadas, elas foram agrupadas em DialogApi 1.2, juntamente com todas as APIs no DialogApi 1.1. *Cada versão de um conjunto de requisitos é um superconjunto de todas as versões anteriores.*

O suporte ao conjunto de requisitos varia de acordo com Office aplicativo, a versão do aplicativo Office e a plataforma na qual ele está sendo executado. Por exemplo, o DialogApi 1.2 não tem suporte em versões de compra avures do Office antes do Office 2021, mas o DialogApi 1.1 tem suporte em todas as versões de compra avures de volta para o Office 2013. Você deseja que seu suplemento seja instalado em cada combinação de plataforma e versão do Office que dá suporte às APIs que ele usa, portanto, você sempre deve especificar no manifesto a versão mínima de cada  conjunto de requisitos que seu suplemento requer. Os detalhes sobre como fazer isso são posteriormente neste artigo.

> [!TIP]
> Para obter mais informações sobre o controle de versão do conjunto de requisitos, consulte [Office](office-versions-and-requirement-sets.md#office-requirement-sets-availability) disponibilidade de conjuntos de requisitos e para obter as listas completas de conjuntos de requisitos e informações sobre as APIs em cada um, comece com Office conjuntos de requisitos de [suplemento.](/javascript/api/requirement-sets/common/office-add-in-requirement-sets) Os tópicos de referência para a maioria Office.js APIs também especificam o conjunto de requisitos ao qual pertencem (se houver).

> [!NOTE]
> Alguns conjuntos de requisitos também têm elementos de manifesto associados a eles. Consulte [Especificando requisitos em um elemento VersionOverrides](#specify-requirements-in-a-versionoverrides-element) para obter informações sobre quando esse fato é relevante para o design do suplemento.

#### <a name="apis-not-in-a-requirement-set"></a>APIs que não estão em um conjunto de requisitos

Todas as APIs nos modelos específicos do aplicativo estão em conjuntos de requisitos, mas algumas delas no modelo de API Comum não estão. Também há uma maneira de especificar uma dessas APIs sem definição no manifesto quando o suplemento exigir uma. Os detalhes estão incluídos mais adiante neste artigo.

### <a name="requirements-element"></a>Elemento Requirements

Use o [elemento Requirements](/javascript/api/manifest/requirements) e seus conjuntos e [](/javascript/api/manifest/sets) métodos de elementos filho para especificar os conjuntos de [requisitos mínimos](/javascript/api/manifest/methods) ou membros da API que devem ter suporte do aplicativo Office para instalar o suplemento. 

Se o aplicativo ou a plataforma do Office não for compatível com os conjuntos de requisitos ou membros da API especificados no elemento **Requirements**, o suplemento não será executado nesse aplicativo ou plataforma e não será exibido em Meus **Suplementos**.

> [!NOTE]
> O **elemento Requirements** é opcional para todos os suplementos, exceto Outlook suplementos. Quando o `xsi:type` atributo do `OfficeApp` `MailApp`elemento raiz é , deve haver um elemento **Requirements** que especifica a versão mínima do conjunto de requisitos de Caixa de Correio que o suplemento requer. Para obter mais informações, consulte [Outlook de requisitos da API JavaScript](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets).

O exemplo de código a seguir mostra como configurar um suplemento que pode ser instalado em todos os Office aplicativos compatíveis com o seguinte:

-  `TableBindings` conjunto de requisitos, que tem uma versão mínima de "1.1".
-  `OOXML` conjunto de requisitos, que tem uma versão mínima de "1.1".
-  `Document.getSelectedDataAsync` Método.

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

- O **elemento Requirements** contém os **elementos** filho **Conjuntos** e Métodos.
- O **elemento Sets** pode conter um ou mais **elementos Set** . `DefaultMinVersion` especifica o valor padrão de `MinVersion` todos os elementos **set** filho.
- Um [elemento Set](/javascript/api/manifest/set) especifica um conjunto de requisitos que o Office aplicativo deve dar suporte para tornar o suplemento inserível. O `Name` atributo especifica o nome do conjunto de requisitos. Especifica `MinVersion` a versão mínima do conjunto de requisitos. `MinVersion` substitui o valor do atributo `DefaultMinVersion` nos Conjuntos **pai**.
- O **elemento Methods** pode conter um ou [mais elementos Method](/javascript/api/manifest/method) . Você não pode usar o elemento **Methods** com suplementos do Outlook.
- O **elemento** Method especifica um método individual que o Office aplicativo deve dar suporte para tornar o suplemento inserível. O `Name` atributo é necessário e especifica o nome do método qualificado com seu objeto pai.

## <a name="design-for-alternate-experiences"></a>Design para experiências alternativas

Os recursos de extensibilidade que a Office de suplementos fornece podem ser divididos de forma útil em três tipos:

- Recursos de extensibilidade que estão disponíveis imediatamente após a instalação do suplemento. Você pode usar esse tipo de recurso configurando um [elemento VersionOverrides](/javascript/api/manifest/versionoverrides) no manifesto. Um exemplo desse tipo de recurso são comandos de [suplemento](../design/add-in-commands.md), que são botões e menus da faixa de opções personalizados.
- Recursos de extensibilidade que estão disponíveis somente quando o suplemento está em execução e que são implementados com Office.js APIs JavaScript; por exemplo, [caixas de diálogo](../develop/dialog-api-in-office-add-ins.md).
- Recursos de extensibilidade que estão disponíveis apenas em runtime, mas são implementados com uma combinação de Office.js JavaScript e configuração em um **elemento VersionOverrides** . Exemplos disso são [Excel funções personalizadas](../excel/custom-functions-overview.md), [logon](sso-in-office-add-ins.md) único e [guias contextuais personalizadas](../design/contextual-tabs.md).

Se o suplemento usa um recurso de extensibilidade específico para algumas de suas funcionalidades, mas tem outras funcionalidades úteis que não exigem o recurso de extensibilidade, você deve projetar o suplemento para que ele seja instalado na plataforma e combinações de versão do Office que não dão suporte ao recurso de extensibilidade. Ele pode fornecer uma experiência valiosa, embora diminuída, nessas combinações. 

Você implementa esse design de maneira diferente, dependendo de como o recurso de extensibilidade é implementado: 

- Para recursos implementados inteiramente com JavaScript, consulte [verificações de runtime para obter suporte ao método e ao conjunto de requisitos](#runtime-checks-for-method-and-requirement-set-support).
- Para recursos que exigem que você configure um **elemento VersionOverrides** , consulte [Especificando requisitos em um elemento VersionOverrides](#specify-requirements-in-a-versionoverrides-element).

### <a name="runtime-checks-for-method-and-requirement-set-support"></a>Verificações de runtime para suporte ao método e ao conjunto de requisitos 

Você testa em runtime para descobrir se o Office do usuário dá suporte a um conjunto de requisitos com o [método isSetSupported](/javascript/api/office/office.requirementsetsupport#office-office-requirementsetsupport-issetsupported-member(1)). Passe o nome do conjunto de requisitos e a versão mínima como parâmetros. Se o conjunto de requisitos for compatível, retornará `isSetSupported` **true**. O código a seguir mostra um exemplo.

```js
if (Office.context.requirements.isSetSupported('WordApi', '1.1'))
{
   // Code that uses API members from the WordApi 1.1 requirement set.
} else {
   // Provide diminished experience here. E.g., run alternate code when the user's Word is one-time purchase Word 2013 (which does not support WordApi 1.1).
}
```
Sobre este código, observe:

- O primeiro parâmetro é necessário. É uma cadeia de caracteres que representa o nome do conjunto de requisitos. Para saber mais sobre os conjuntos de requisitos disponíveis, confira [Conjuntos de requisitos de Suplemento do Office](/javascript/api/requirement-sets/common/office-add-in-requirement-sets).
- O segundo parâmetro é opcional. É uma cadeia de caracteres que especifica a versão mínima do conjunto de requisitos que o aplicativo Office `if` deve dar suporte para que o código dentro da instrução seja executado (por exemplo, "**1.9**"). Se não for usada, a versão "1.1" será assumida.

> [!WARNING]
> Ao chamar o `isSetSupported` método, o valor do segundo parâmetro (se especificado) deve ser uma cadeia de caracteres e não um número. O analisador JavaScript não pode diferenciar entre valores numéricos como 1.1 e 1.10, enquanto pode para valores de cadeia de caracteres como "1.1" e "1.10".

A tabela a seguir mostra os nomes do conjunto de requisitos para os modelos de API específicos do aplicativo.

|Aplicativo do Office|RequirementSetName|
|---|---|
|Excel|ExcelApi|
|OneNote|OneNoteApi|
|Outlook|Caixa de correio|
|PowerPoint|PowerPointApi|
|Word|WordApi|

A seguir está um exemplo de como usar o método com um dos conjuntos de requisitos do modelo de API comum.

```js
if (Office.context.requirements.isSetSupported('CustomXmlParts'))
{
    // Run code that uses API members from the CustomXmlParts requirement set.
}
else
{
    // Run alternate code when the user's Word doesn't support the CustomXmlParts requirement set.
}
```

> [!NOTE] 
> O `isSetSupported` método e os conjuntos de requisitos para esses aplicativos estão disponíveis no arquivo Office.js mais recente no CDN. Se você não usar o Office.js do CDN, `isSetSupported` seu suplemento poderá gerar exceções se você estiver usando uma versão antiga da biblioteca na qual é indefinida. Para obter mais informações, [consulte Usar a biblioteca Office API JavaScript mais recente](#use-the-latest-office-javascript-api-library).

Quando o suplemento depende de um método que não faz parte de um conjunto de requisitos, use a verificação de runtime para determinar se o método é compatível com o aplicativo Office, conforme mostrado no exemplo de código a seguir. Para obter uma lista completa dos métodos que não pertencem a um conjunto de requisitos, confira [Conjuntos de requisitos de Suplemento do Office](/javascript/api/requirement-sets/common/office-add-in-requirement-sets#methods-that-arent-part-of-a-requirement-set).

> [!NOTE]
> Recomendamos limitar o uso desse tipo de verificação no tempo de execução no código de seu suplemento.

O exemplo de código a seguir verifica se o aplicativo Office dá suporte`document.setSelectedDataAsync`.

```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses `document.setSelectedDataAsync`.
}
```

### <a name="specify-requirements-in-a-versionoverrides-element"></a>Especificar requisitos em um elemento VersionOverrides

O elemento [VersionOverrides](/javascript/api/manifest/versionoverrides) foi adicionado ao esquema de manifesto principalmente, mas não exclusivamente, para dar suporte a recursos que devem estar disponíveis imediatamente após a instalação de um suplemento, como comandos de suplemento (botões e menus da faixa de opções personalizados). Office deve saber sobre esses recursos quando analisa o manifesto do suplemento. 

Suponha que seu suplemento use um desses recursos, mas o suplemento é valioso e deve ser instalado, mesmo em versões Office que não dão suporte ao recurso. Nesse cenário, identifique o recurso usando um elemento [Requirements](/javascript/api/manifest/requirements) (e seus elementos Filho [](/javascript/api/manifest/sets) conjuntos e [](/javascript/api/manifest/methods) métodos) que você inclui como um filho do próprio elemento **VersionOverrides** em vez de como um filho do elemento base`OfficeApp`. O efeito de fazer isso é que o Office permitirá que o suplemento seja instalado, mas o Office ignorará alguns dos elementos filho do **elemento VersionOverrides** em versões Office em que não há suporte para o recurso.

Especificamente, os elementos filho dos **VersionOverrides** que substituem elementos no manifesto base, como um elemento **Hosts** , são ignorados e os elementos correspondentes do manifesto base são usados em vez disso. No entanto, pode haver elementos filho em um **VersionOverrides** que implementam recursos adicionais em vez de substituir as configurações no manifesto base. Dois exemplos são o `WebApplicationInfo` e `EquivalentAddins`. Essas partes do **VersionOverrides** não serão  ignoradas, supondo que a plataforma e a versão Office suporte ao recurso correspondente.  

Para obter informações sobre os elementos descendentes do **elemento Requirements** , consulte [o elemento Requirements](#requirements-element) anteriormente neste artigo.

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
> Tenha muito cuidado antes de usar um elemento **Requirements** em um **VersionOverrides**, pois em combinações de plataforma e versão que não dão suporte ao *requisito, nenhum* dos comandos de suplemento será *instalado, mesmo* aqueles que invocam a funcionalidade que não precisa do requisito. Considere, por exemplo, um suplemento que tenha dois botões de faixa de opções personalizados. Uma delas chama Office APIs JavaScript disponíveis no conjunto de requisitos **ExcelApi 1.4** (e posterior). As outras chamadas apIs que só estão disponíveis no **ExcelApi 1.9** (e posterior). Se você colocar um requisito para **o ExcelApi 1.9** no **VersionOverrides**, quando não houver suporte para 1.9, nenhum botão será exibido  na faixa de opções. Uma estratégia melhor nesse cenário seria usar a técnica descrita nas verificações de [runtime para suporte ao método e ao conjunto de requisitos](#runtime-checks-for-method-and-requirement-set-support). O código invocado pelo segundo botão primeiro usa `isSetSupported` para verificar se há suporte do **ExcelApi 1.9**. Se não houver suporte, o código fornecerá ao usuário uma mensagem informando que esse recurso do suplemento não está disponível em sua versão do Office. 

> [!TIP]
> Não há nenhum ponto para repetir um elemento **Requirement** em um **VersionOverrides** que já aparece no manifesto base. Se o requisito for especificado no manifesto base, o suplemento não poderá instalar onde não há suporte para o requisito, portanto, Office nem mesmo analisa o elemento **VersionOverrides**. 

## <a name="see-also"></a>Confira também

- [Manifesto XML dos Suplementos do Office](add-in-manifests.md)
- [Conjuntos de requisitos de Suplemento do Office](/javascript/api/requirement-sets/common/office-add-in-requirement-sets)
- [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)
