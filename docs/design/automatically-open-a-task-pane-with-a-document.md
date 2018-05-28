---
title: Abrir automaticamente um painel de tarefas com um documento
description: ''
ms.date: 05/02/2018
ms.openlocfilehash: 06e1cce3a45a5af744a1be4b3feabbf051940d76
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="automatically-open-a-task-pane-with-a-document"></a>Abrir automaticamente um painel de tarefas com um documento

Voc? pode usar comandos de suplemento no seu Suplemento do Office para estender a interface do usu?rio do Office adicionando bot?es ? faixa de op??es do Office. Quando os usu?rios clicam no bot?o de comando, ocorre uma a??o, como abrir um painel de tarefas. 

Alguns cen?rios exigem que um painel de tarefas abra automaticamente quando um documento ? aberto, sem a intera??o expl?cita do usu?rio. Voc? pode usar o recurso autoopen do painel de tarefas, apresentado no conjunto de requisitos AddInCommands 1.1, para abrir automaticamente um painel de tarefas quando necess?rio. 


## <a name="how-is-the-autoopen-feature-different-from-inserting-a-task-pane"></a>De que forma o recurso autoopen ? diferente da inser??o de um painel de tarefas? 

Quando um usu?rio lan?ar suplementos que n?o usam comandos de suplemento, por exemplo, suplementos que s?o executados no Office 2013, eles ser?o inseridos no documento e persistir?o nesse documento. Como resultado, quando outros usu?rios abrem o documento, ? solicitado que eles instalem o suplemento, e o painel de tarefas abrir?. O desafio com esse modelo ? que, em muitos casos, os usu?rios n?o querem que o suplemento persista no documento. Por exemplo, um aluno que usa um suplemento de dicion?rio em um documento do Word pode n?o querer que seus colegas ou professores sejam avisados para instalar esse suplemento quando abrirem o documento.  

Com o recurso autoopen, voc? pode explicitamente definir, ou permitir que o usu?rio defina, se um suplemento do painel de tarefas ir? persistir em um documento espec?fico. 

## <a name="support-and-availability"></a>Suporte e disponibilidade
O recurso autoopen atualmente tem suporte do <!-- in **developer preview** and it is only --> nos seguintes produtos e plataformas.

|**Produtos**|**Plataformas**|
|:-----------|:------------|
|<ul><li>Word</li><li>Excel</li><li>PowerPoint</li></ul>|Plataformas suportadas para todos os produtos:<ul><li>Office para Windows Desktop. Vers?o 16.0.8121.1000+</li><li>Office para Mac. Vers?o 15.34.17051500+</li><li>Office Online</li></ul>|


## <a name="best-practices"></a>Pr?ticas recomendadas

Aplique as seguintes pr?ticas recomendadas ao usar o recurso autoopen:

- Use o recurso autoopen quando ele auxiliar a efici?ncia dos usu?rios do seu suplemento, como
    - Quando o documento precisa do suplemento para funcionar corretamente. Por exemplo, uma planilha que inclui valores de a??es que s?o atualizados periodicamente por um suplemento. O suplemento dever? abrir automaticamente quando a planilha for aberta para manter os valores atualizados. 
    - Quando ? muito prov?vel que o usu?rio sempre utilizar? o suplemento com um determinado documento. Por exemplo, um suplemento que ajuda os usu?rios a preencher ou alterar dados em um documento puxando informa??es de um sistema de back-end. 
- Permita que os usu?rios ativem ou desativem o recurso autoopen. Inclua uma op??o em sua interface de usu?rio para que eles possam escolher quando n?o querem mais que o suplemento abra automaticamente no painel de tarefas.  
- Use a detec??o de configura??o de exig?ncia para determinar se o recurso autoopen est? dispon?vel e fornecer um comportamento de fallback se ele n?o estiver dispon?vel.
- N?o use o recurso autoopen para aumentar artificialmente o uso do seu suplemento. Se n?o faz sentido seu suplemento abrir automaticamente em determinados documentos, esse recurso pode incomodar os usu?rios. 

    > [!NOTE]
    > Se a Microsoft detectar abuso do recurso autoopen, seu suplemento poder? ser rejeitado no AppSource. 

- N?o use esse recurso para fixar v?rios pain?is de tarefas. Voc? s? pode definir um painel do suplemento para abrir automaticamente com um documento.  

## <a name="implementation"></a>Implementa??o
Para implementar o recurso autoopen:

- Especifique o painel de tarefas a ser aberto automaticamente.
- Marque o documento para abrir o painel de tarefas automaticamente.

> [!IMPORTANT]
> O painel que voc? designar para abrir automaticamente s? ser? aberto se o suplemento j? estiver instalado no dispositivo do usu?rio. Se o usu?rio n?o tiver o suplemento instalado quando abrir um documento, o recurso autoopen n?o funcionar?, e a configura??o ser? ignorada. Se voc? tamb?m exigir que o suplemento seja distribu?do com o documento, ser? preciso definir a propriedade de visibilidade como 1. Isso s? pode ser feito usando OpenXML. Um exemplo ser? fornecido posteriormente neste artigo. 

### <a name="step-1-specify-the-task-pane-to-open"></a>Etapa 1: especificar o painel de tarefas que ser? aberto
Para especificar o painel de tarefas que ser? aberto automaticamente, defina o valor [TaskpaneId](https://dev.office.com/reference/add-ins/manifest/action#taskpaneid) para **Office.AutoShowTaskpaneWithDocument**. Voc? s? pode definir esse valor em um painel de tarefas. Se voc? definir esse valor em v?rios pain?is de tarefas, a primeira ocorr?ncia do valor ser? reconhecida e as outras ser?o ignoradas. 

O exemplo a seguir mostra o valor TaskPaneId configurado para Office.AutoShowTaskpaneWithDocument.
          
```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="Contoso.Taskpane.Url" />
</Action>
```     

### <a name="step-2-tag-the-document-to-automatically-open-the-task-pane"></a>Etapa 2: marcar o documento para abrir o painel de tarefas automaticamente

Voc? pode marcar o documento para acionar o recurso autoopen de duas maneiras. Escolha a alternativa que funciona melhor para o seu cen?rio.  


#### <a name="tag-the-document-on-the-client-side"></a>Marcar o documento no lado do cliente
Use o m?todo [settings.set](https://dev.office.com/reference/add-ins/shared/settings.set) do Office.js para configurar o **Office.AutoShowTaskpaneWithDocument** para **true**, conforme mostrado no exemplo a seguir.   

```js
Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
Office.context.document.settings.saveAsync();
```

Use esse m?todo se voc? precisar marcar o documento como parte da intera??o com o suplemento (por exemplo, assim que o usu?rio criar uma liga??o ou escolher uma op??o para indicar que deseja que o painel abra automaticamente).

#### <a name="use-open-xml-to-tag-the-document"></a>Usar Open XML para marcar o documento
Voc? pode usar o Open XML para criar ou modificar um documento e adicionar a marca??o XML do Open Office apropriada para acionar o recurso autoopen. Veja um exemplo de como fazer isso em [Office-OOXML-EmbedAddin](https://github.com/OfficeDev/Office-OOXML-EmbedAddin). 

Adicione duas partes do Open XML no documento:

- Uma parte webextension
- Uma parte do painel de tarefas

O exemplo a seguir mostra como adicionar a parte webextension.

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

A parte webextension inclui um conjunto de propriedades e uma propriedade chamada **Office.AutoShowTaskpaneWithDocument** que deve ser definida para `true`.

A parte webextension tamb?m inclui uma refer?ncia para a loja ou o cat?logo com atributos para `id`, `storeType`, `store` e `version`. Do valores `storeType`, somente quatro s?o relevantes para o recurso autoopen. Os valores dos outros tr?s atributos dependem do valor de `storeType`, conforme mostrado na tabela a seguir. 

| **`storeType` valor** | **`id` valor**    |**`store` valor** | **`version` valor**|
|:---------------|:---------------|:---------------|:---------------|
|OMEX (AppSource)|A ID do ativo do suplemento no AppSource (confira a observa??o)|A localidade do AppSource, por exemplo, "pt-br".|A vers?o no cat?logo do AppSource (confira a observa??o)|
|FileSystem (um compartilhamento de rede)|O GUID do suplemento no manifesto do suplemento.|O caminho do compartilhamento de rede. Por exemplo, "\\\\Meu Computador\\Minha Pasta Compartilhada".|A vers?o no manifesto do suplemento.|
|EXCatalog (implanta??o por meio do servidor Exchange) |O GUID do suplemento no manifesto do suplemento.|"EXCatalog". A linha EXCatalog ? a linha a ser usada com suplementos que usam a Implanta??o Centralizada no Centro de administra??o do Office 365.|A vers?o no manifesto do suplemento.
|Registro (registro de sistema)|O GUID do suplemento no manifesto do suplemento.|"desenvolvedor"|A vers?o no manifesto do suplemento.|

> [!NOTE]
> Para localizar a ID de ativos e a vers?o de um suplemento no AppSource, v? para a p?gina inicial do suplemento no AppSource. A ID de ativo aparece na barra de endere?os no navegador. A vers?o aparece na se??o **Detalhes** da p?gina.

Confira mais informa??es sobre a marca??o webextension em [[MS-OWEXML] 2.2.5. WebExtensionReference](https://msdn.microsoft.com/en-us/library/hh695383(v=office.12).aspx).

O exemplo a seguir mostra como adicionar a parte do painel de tarefas.

```xml
<wetp:taskpane dockstate="right" visibility="0" width="350" row="4" xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
  <wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" />
</wetp:taskpane>
```

Observe que neste exemplo, o atributo `visibility` est? definido como "0". Isso significa que ap?s serem adicionadas as partes webextension e de painel de tarefas, a primeira vez que o documento for aberto, o usu?rio dever? instalar o suplemento clicando no bot?o **Suplemento** na faixa de op??es. Depois disso, o painel de tarefas do suplemento abrir? automaticamente quando o arquivo for aberto. Al?m disso, ao definir `visibility` como "0", ? poss?vel usar o Office.js para permitir que os usu?rios ativem ou desativem o recurso autoopen. Especificamente, seu script define a configura??o de documento **Office.AutoShowTaskpaneWithDocument** para `true` ou `false`. Confira mais detalhes em [Marcar o documento no lado do cliente](#tag-the-document-on-the-client-side). 

Se o elemento `visibility` ? definido como "1", o painel de tarefas abrir? automaticamente na primeira vez em que o documento for aberto. O usu?rio ? solicitado a confiar no suplemento e, quando a confian?a ? concedida, o suplemento ? aberto. Depois disso, o painel de tarefas do suplemento abrir? automaticamente quando o arquivo for aberto. Entretanto, ao definir `visibility` como "1", n?o ? poss?vel usar o Office.js para permitir que os usu?rios ativem ou desativem o recurso autoopen. 

Definir o `visibility` como "1" ? uma boa op??o quando o suplemento e o modelo ou o conte?do do documento s?o muito estreitamente integrados de modo que o usu?rio n?o poderia optar por cancelar o recurso autoopen. 

> [!NOTE]
> Se quiser distribuir seu suplemento com o documento, para que os usu?rios sejam solicitados a instal?-lo, voc? dever? definir a propriedade de visibilidade para 1. Isso s? pode ser feito pelo Open XML.

Uma maneira f?cil de escrever o XML ? primeiro executar seu suplemento e [marcar o documento no lado do cliente](#tag-the-document-on-the-client-side) para escrever o valor e, em seguida, salvar o documento e inspecionar o XML que ? gerado. O Office detectar? e fornecer? os valores de atributo apropriados. Voc? tamb?m pode usar a [Ferramenta de Produtividade Open XML SDK 2.5](https://www.microsoft.com/en-us/download/details.aspx?id=30425) para gerar o c?digo C# para adicionar por meio de programa??o a marca??o com base no XML que voc? gerou.

## <a name="test-and-verify-opening-taskpanes"></a>Teste e verifique a abertura dos pain?is de tarefas
Voc? pode implantar uma vers?o de teste do seu suplemento que abrir? automaticamente um painel de tarefas usando a Implanta??o Centralizada atrav?s do Centro de administra??o do Office 365. O exemplo a seguir mostra como os suplementos s?o inseridos a partir do cat?logo de Implanta??o Centralizada usando a vers?o da loja do EXCatalog.

```xml
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{52811C31-4593-43B8-A697-EB873422D156}">
    <we:reference id="af8fa5ba-4010-4bcc-9e03-a91ddadf6dd3" version="1.0.0.0" store="EXCatalog" storeType="EXCatalog"/>
    <we:alternateReferences/>
    <we:properties/>
    <we:bindings/>
    <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
```
Para testar o exemplo anterior, consulte [Configurar seu ambiente de desenvolvimento do Office 365](https://msdn.microsoft.com/en-us/office/office365/howto/setup-development-environment) e considere assinar uma [Conta para desenvolvedores do Office 365](https://developer.microsoft.com/en-us/office/dev-program). Voc? pode fazer um test drive da Implanta??o Centralizada e verificar se o suplemento funciona conforme o esperado.


## <a name="see-also"></a>Veja tamb?m

Para saber como usar o recurso autoopen, confira as [Amostras de comandos do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/AutoOpenTaskpane). 
[Junte-se ao programa para desenvolvedores do Office 365](https://docs.microsoft.com/en-us/office/developer-program/office-365-developer-program). 

