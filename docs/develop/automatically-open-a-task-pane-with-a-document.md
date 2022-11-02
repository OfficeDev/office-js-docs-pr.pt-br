---
title: Abrir automaticamente um painel de tarefas com um documento
description: Saiba como configurar um Suplemento do Office para abrir automaticamente quando um documento for aberto.
ms.date: 09/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: 125e6bcccceb9fe0ced6992ba04a954695235ed4
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810187"
---
# <a name="automatically-open-a-task-pane-with-a-document"></a>Abrir automaticamente um painel de tarefas com um documento

Você pode usar comandos de suplemento no suplemento do Office para estender a interface do usuário do Office adicionando botões à faixa de opções do aplicativo do Office. Quando os usuários clicam no botão de comando, ocorre uma ação, como abrir um painel de tarefas.

Alguns cenários exigem que um painel de tarefas seja exibido automaticamente ao abrir um documento, sem a interação explícita do usuário. Você pode usar o recurso de painel de tarefas de abertura automática, introduzido no [conjunto de requisitos AddInCommands 1.1](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets), para abrir automaticamente um painel de tarefas quando o cenário exigir.

> [!NOTE]
> Para configurar um painel de tarefas a ser aberto imediatamente quando o suplemento estiver instalado, mas não necessariamente quando o documento for aberto posteriormente, consulte [Abrir automaticamente um painel de tarefas quando um suplemento estiver instalado](automatically-open-on-installation.md).

## <a name="how-is-the-autoopen-feature-different-from-inserting-a-task-pane"></a>De que forma o recurso autoopen é diferente da inserção de um painel de tarefas?

Quando um usuário lançar suplementos que não usam comandos de suplemento, por exemplo, suplementos que são executados no Office 2013, eles serão inseridos no documento e persistirão nesse documento. Como resultado, quando outros usuários abrem o documento, é solicitado que eles instalem o suplemento, e o painel de tarefas abrirá. O desafio com esse modelo é que, em muitos casos, os usuários não querem que o suplemento persista no documento. Por exemplo, um aluno que usa um suplemento de dicionário em um documento do Word pode não querer que seus colegas ou professores sejam avisados para instalar esse suplemento quando abrirem o documento.

Com o recurso autoopen, você pode explicitamente definir, ou permitir que o usuário defina, se um suplemento do painel de tarefas irá persistir em um documento específico.

## <a name="support-and-availability"></a>Suporte e disponibilidade

Atualmente, o recurso de preenchimento automático tem suporte nos seguintes produtos e plataformas.

|Produtos|Plataformas|
|:-----------|:------------|
|<ul><li>Word</li><li>Excel</li><li>PowerPoint</li></ul>|Plataformas com suporte para todos os produtos com suporte:<ul><li>Office na Área de Trabalho do Windows. Versão 16.0.8121.1000+</li><li>Office on Mac. Build 15.34.17051500+</li><li>Office na Web</li></ul>|

## <a name="best-practices"></a>Práticas recomendadas

Aplique as práticas recomendadas a seguir ao usar o recurso de preenchimento automático.

- Use o recurso autoopen quando ele auxiliar a eficiência dos usuários do seu suplemento, como
  - When the document needs the add-in in order to function properly. For example, a spreadsheet that includes stock values that are periodically refreshed by an add-in. The add-in should open automatically when the spreadsheet is opened to keep the values up to date.
  - When the user will most likely always use the add-in with a particular document. For example, an add-in that helps users fill in or change data in a document by pulling information from a backend system.
- Allow users to turn on or turn off the autoopen feature. Include an option in your UI for users to choose to no longer automatically open the add-in task pane.  
- Use a detecção de conjunto de requisitos para determinar se o recurso de preenchimento automático está disponível e fornecer um comportamento de fallback se não estiver.
- Não use o recurso autoopen para aumentar artificialmente o uso do seu suplemento. Se não fizer sentido que seu suplemento seja aberto automaticamente com determinados documentos, esse recurso poderá incomodar os usuários.

    > [!NOTE]
    > Se a Microsoft detectar abuso do recurso autoopen, seu suplemento poderá ser rejeitado no AppSource.

- Don't use this feature to pin multiple task panes. You can only set one pane of your add-in to open automatically with a document.  

## <a name="implement-the-autoopen-feature"></a>Implementar o recurso de preenchimento automático

- Especifique o painel de tarefas a ser aberto automaticamente.
- Marque o documento para abrir o painel de tarefas automaticamente.

> [!IMPORTANT]
> The pane that you designate to open automatically will only open if the add-in is already installed on the user's device. If the user does not have the add-in installed when they open a document, the autoopen feature will not work and the setting will be ignored. If you also require the add-in to be distributed with the document you need to set the visibility property to 1; this can only be done using OpenXML, an example is provided later in this article.

### <a name="step-1-specify-the-task-pane-to-open"></a>Etapa 1: especificar o painel de tarefas que será aberto

To specify the task pane to open automatically, set the [TaskpaneId](/javascript/api/manifest/action#taskpaneid) value to **Office.AutoShowTaskpaneWithDocument**. You can only set this value on one task pane. If you set this value on multiple task panes, the first occurrence of the value will be recognized and the others will be ignored.

O exemplo a seguir mostra o valor TaskPaneId configurado para Office.AutoShowTaskpaneWithDocument.

```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="Contoso.Taskpane.Url" />
</Action>
```

### <a name="step-2-tag-the-document-to-automatically-open-the-task-pane"></a>Etapa 2: marcar o documento para abrir o painel de tarefas automaticamente

You can tag the document to trigger the autoopen feature in one of two ways. Pick the alternative that works best for your scenario.  

#### <a name="tag-the-document-on-the-client-side"></a>Marcar o documento no lado do cliente

Use o método Office.js [settings.set](/javascript/api/office/office.settings) para definir **Office.AutoShowTaskpaneWithDocument** `true`como , conforme mostrado no exemplo a seguir.

```js
Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
Office.context.document.settings.saveAsync();
```

Use esse método se você precisar marcar o documento como parte da interação com o suplemento (por exemplo, assim que o usuário criar uma ligação ou escolher uma opção para indicar que deseja que o painel abra automaticamente).

#### <a name="use-open-xml-to-tag-the-document"></a>Usar Open XML para marcar o documento

You can use Open XML to create or modify a document and add the appropriate Open Office XML markup to trigger the autoopen feature. For a sample that shows you how to do this, see [Office-OOXML-EmbedAddin](https://github.com/OfficeDev/Office-OOXML-EmbedAddin).

Adicione duas partes XML abertas ao documento.

- Uma parte `webextension`
- Uma parte `taskpane`

O exemplo a seguir mostra como adicionar a parte `webextension`.

```xml
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="[ADD-IN ID PER MANIFEST]">
  <we:reference id="[GUID or AppSource asset ID]" version="[your add-in version]" store="[Pointer to store or catalog]" storeType="[Store or catalog type]"/>
  <we:alternateReferences/>
  <we:properties>
   <we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/>
  </we:properties>
  <we:bindings/>
  <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
```

A parte `webextension` inclui um conjunto de propriedades e uma propriedade chamada **Office.AutoShowTaskpaneWithDocument** que deve ser definida como `true`.

A parte `webextension` também inclui uma referência para a loja ou o catálogo com atributos para `id`, `storeType`, `store` e `version`. Dos valores `storeType`, somente quatro são relevantes para o recurso autoopen. Os valores dos outros três atributos dependem do valor de `storeType`, conforme mostrado na tabela a seguir.

|`storeType` valor|`id` valor|`store` valor|`version` valor|
|:---------------|:---------------|:---------------|:---------------|
|OMEX (AppSource)|A ID do ativo AppSource do suplemento (consulte Observação).|A localidade do AppSource, por exemplo, "pt-br".|A versão no catálogo do AppSource (consulte Observação).|
|WOPICatalog (hosts [WOPI](/microsoft-365/cloud-storage-partner-program/online/) de parceiro)| A ID do ativo AppSource do suplemento (consulte Observação). | "wopicatalog". Use esse valor para suplementos publicados na Fonte do Aplicativo e instalados em hosts WOPI. Para obter mais informações, confira [Integrando-se ao Office Online](/microsoft-365/cloud-storage-partner-program/online/overview). | A versão no manifesto do suplemento.|
|FileSystem (um compartilhamento de rede)|O GUID do suplemento no manifesto do suplemento.|O caminho do compartilhamento de rede. Por exemplo, "\\\\Meu Computador\\Minha Pasta Compartilhada".|A versão no manifesto do suplemento.|
|EXCatalog (implantação por meio do servidor Exchange) |O GUID do suplemento no manifesto do suplemento.|"EXCatalog". A linha EXCatalog é a linha a ser usada com suplementos que usam a Implantação Centralizada no Centro de administração do Microsoft 365.|A versão no manifesto do suplemento.|
|Registro (registro de sistema)|O GUID do suplemento no manifesto do suplemento.|"developer"|A versão no manifesto do suplemento.|

> [!NOTE]
> To find the asset ID and version of an add-in in AppSource, go to the AppSource landing page for the add-in. The asset ID appears in the address bar in the browser. The version is listed in the **Details** section of the page.

Saiba mais sobre a marcação webextension em [[MS-OWEXML] 2.2.5. WebExtensionReference](/openspecs/office_standards/ms-owexml/d4081e0b-5711-45de-b708-1dfa1b943ad1).

O exemplo a seguir mostra como adicionar a parte `taskpane`.

```xml
<wetp:taskpane dockstate="right" visibility="0" width="350" row="4" xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
  <wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" />
</wetp:taskpane>
```

Observe que neste exemplo, o atributo `visibility` está definido como "0". Isso significa que, após adicionar as partes webextension e `taskpane`, a primeira vez que o documento for aberto, o usuário deve instalar o suplemento clicando no botão **Suplemento** na faixa de opções. Depois disso, o painel de tarefas do suplemento abre automaticamente quando o arquivo for aberto. E, ao definir `visibility` como "0", é possível usar o Office.js para permitir que os usuários ativem ou desativem o recurso autoopen. Especificamente, seu script define a configuração de documento **Office.AutoShowTaskpaneWithDocument** como `true` ou `false`. (Saiba mais em [Marcar o documento no lado do cliente](#tag-the-document-on-the-client-side).)

If `visibility` is set to "1", the task pane opens automatically the first time the document is opened. The user is prompted to trust the add-in, and when trust is granted, the add-in opens. Thereafter, the add-in task pane opens automatically when the file is opened. However, when `visibility` is set to "1", you can't use Office.js to enable users to turn on or turn off the autoopen feature.

Definir o `visibility` como "1" é uma boa opção quando o suplemento e o modelo ou o conteúdo do documento são muito estreitamente integrados de modo que o usuário não poderia optar por cancelar o recurso autoopen.

> [!NOTE]
> If you want to distribute your add-in with the document, so that users are prompted to install it, you must set the visibility property to 1. You can only do this via Open XML.

Uma maneira fácil de gravar o XML é primeiro executar o suplemento e [marcar o documento no lado do cliente](#tag-the-document-on-the-client-side) para gravar o valor e, em seguida, salvar o documento e inspecionar o XML gerado. O Office detectará e fornecerá os valores de atributo apropriados. Você também pode usar a [Ferramenta de Produtividade do SDK Open XML](https://www.nuget.org/packages/Open-XML-SDK) para gerar código C# para adicionar programaticamente a marcação com base no XML gerado.

## <a name="test-and-verify-opening-task-panes"></a>Testar e verificar a abertura de painéis de tarefas

Você pode implantar uma versão de teste do suplemento que abrirá automaticamente um painel de tarefas usando a Implantação Centralizada por meio do Centro de administração do Microsoft 365. O exemplo a seguir mostra como os suplementos são inseridos do catálogo de Implantação Centralizada usando a versão de armazenamento EXCatalog.

```xml
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{52811C31-4593-43B8-A697-EB873422D156}">
    <we:reference id="af8fa5ba-4010-4bcc-9e03-a91ddadf6dd3" version="1.0.0.0" store="EXCatalog" storeType="EXCatalog"/>
    <we:alternateReferences/>
    <we:properties/>
    <we:bindings/>
    <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
```

Você pode testar o exemplo anterior usando sua assinatura do Microsoft 365 para experimentar a Implantação Centralizada e verificar se o suplemento funciona conforme o esperado. Se você ainda não tiver uma assinatura do Microsoft 365, poderá obter uma assinatura gratuita e renovável do Microsoft 365 de 90 dias ingressando no [programa de desenvolvedor do Microsoft 365](https://developer.microsoft.com/office/dev-program).

## <a name="see-also"></a>Confira também

- Para ver um exemplo que mostra como usar o recurso autoopen, consulte os [exemplos de comandos do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/AutoOpenTaskpane).
- [Abra automaticamente um painel de tarefas quando um suplemento é instalado](automatically-open-on-installation.md)
- [Ingresse no programa de desenvolvedor do Microsoft 365.](/office/developer-program/office-365-developer-program)