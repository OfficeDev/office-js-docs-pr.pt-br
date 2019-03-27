---
title: Elemento Action no arquivo de manifesto
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 59df6cce6af1277f365a1dd3cd0b3ef11230804e
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870825"
---
# <a name="action-element"></a>Elemento Action

Especifica a ação a ser executada quando o usuário seleciona controles de [Button](control.md#button-control) ou [Menu](control.md#menu-dropdown-button-controls).

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Sim  | Tipo de ação a executar|

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Descrição  |
|:-----|:-----|
|  [FunctionName](#functionname) |    Especifica o nome da função a executar. |
|  [SourceLocation](#sourcelocation) |    Especifica o local do arquivo de origem para essa ação. |
|  [TaskpaneId](#taskpaneid) | Especifica a ID do contêiner do painel de tarefas.|
|  [Title](#title) | Especifica o título personalizado do painel de tarefas.|
|  [SupportsPinning](#supportspinning) | Especifica se um painel de tarefas tem suporte para fixação, que mantém o painel de tarefas aberto quando o usuário altera a seleção.|
  

## <a name="xsitype"></a>xsi:type

Este atributo especifica o tipo de ação realizada quando o usuário seleciona o botão. Pode ser uma das seguintes:

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a>FunctionName

Elemento obrigatório quando **xsi:type** é "ExecuteFunction". Especifica o nome da função a ser executada. A função está contida no arquivo especificado no elemento [FunctionFile](functionfile.md).

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a>SourceLocation

Elemento obrigatório quando **xsi:type** for "ShowTaskpane". Especifica o local do arquivo de origem para essa ação. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **Url** no elemento **Urls** do elemento [Resources](resources.md).

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a>TaskpaneId

Elemento opcional quando  **xsi:type** for "ShowTaskpane". Especifica a ID do contêiner do painel de tarefas. Quando você tiver várias ações "ShowTaskpane", use uma **TaskpaneId** diferente se desejar ter um painel independente para cada uma. Use a mesma **TaskpaneId** para diferentes ações que compartilhem o mesmo painel. Quando os usuários escolhem comandos que compartilham a mesma **TaskpaneId**, o contêiner do painel permanece aberto, mas o conteúdo do painel é substituído pela ação "SourceLocation" correspondente.

> [!NOTE]
> Esse elemento não tem suporte no Outlook.

O exemplo a seguir mostra duas ações que compartilham o mesmo **TaskpaneId**.

```xml
<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="aTaskPaneUrl" />
</Action>

<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="anotherTaskPaneUrl" />
</Action>
```  

O exemplo a seguir mostra duas ações que usam um **TaskpaneId** diferente. Para ver esses exemplos em contexto, consulte [Exemplo de comando de suplemento simples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).

```xml
<Action xsi:type="ShowTaskpane">
   <TaskpaneId>MyTaskPaneID1</TaskpaneId>
   <SourceLocation resid="Contoso.Taskpane1.Url" />
</Action>

<Action xsi:type="ShowTaskpane">
   <TaskpaneId>MyTaskPaneID2</TaskpaneId>
   <SourceLocation resid="Contoso.Taskpane2.Url" />
</Action>
```  

```xml
<bt:Urls>
   <bt:Url id="Contoso.Taskpane1.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
   <bt:Url id="Contoso.Taskpane2.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane2.html" />
</bt:Urls>
```  

## <a name="title"></a>Cargo

Elemento opcional quando  **xsi:type** for "ShowTaskpane". Especifica o título personalizado do painel de tarefas desta ação.

Os exemplos a seguir mostram duas ações distintas que usam o elemento **Title**.

```xml
<Action xsi:type="ShowTaskpane">
<TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
<SourceLocation resid="PG.Code.Url" />
<Title resid="PG.CodeCommand.Title" />
</Action>
```

```xml
<Action xsi:type="ShowTaskpane">
<SourceLocation resid="PG.Run.Url" />
<Title resid="PG.RunCommand.Title" />
</Action>
```

```xml
<bt:Urls>
<bt:Url id="PG.Code.Url" DefaultValue="https://localhost:3000?commands=1" />
<bt:Url id="PG.Run.Url" DefaultValue="https://localhost:3000/run.html" />
</bt:Urls>
```

```xml
<bt:ShortStrings>
<bt:String id="PG.CodeCommand.Title" DefaultValue="Code" />
<bt:String id="PG.RunCommand.Title" DefaultValue="Run" />
</bt:ShortStrings>
```

## <a name="supportspinning"></a>SupportsPinning

Elemento opcional quando **xsi:type** for "ShowTaskpane". Os elementos [VersionOverrides](versionoverrides.md) incluídos devem ter um valor `VersionOverridesV1_1` para o atributo `xsi:type`. Inclua esse elemento com um valor `true` a fim de fornecer suporte para fixação do painel de tarefas. O usuário pode "fixar" o painel de tarefas, fazendo com que ele permaneça aberto quando alterar a seleção. Para saber mais, consulte [Implementar um painel de tarefas fixável no Outlook](/outlook/add-ins/pinnable-taskpane).

> [!NOTE]
> No momento, o SupportsPinning só tem suporte no Outlook 2016 para Windows (Build 7628,1000 ou posterior) e Outlook 2016 para Mac (Build 16.13.503 ou posterior).

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
