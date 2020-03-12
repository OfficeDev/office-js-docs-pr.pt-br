---
title: Elemento Action no arquivo de manifesto
description: ''
ms.date: 02/28/2020
localization_priority: Normal
ms.openlocfilehash: f7bd577fea1672f592f2b1bac2823d96f0e8a134
ms.sourcegitcommit: 6c7c98f085dd20f827e0c388e672993412944851
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/06/2020
ms.locfileid: "42554907"
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

Elemento obrigatório quando **xsi: Type** for "ShowTaskpane". Especifica o local do arquivo de origem para essa ação. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **Url** no elemento **Urls** do elemento [Resources](resources.md).

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

O exemplo a seguir mostra uma ação que usa o elemento **title** . Observe que você não atribui o **título** a uma cadeia de caracteres diretamente. Em vez disso, atribua um ID de recurso (Resid), que é definido na seção **recursos** do manifesto.

```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="PG.Code.Url" />
    <Title resid="PG.CodeCommand.Title" />
</Action>

 ... Other markup omitted ...
<Resources>
    <bt:Images> ...
    </bt:Images>
    <bt:Urls>
        <bt:Url id="PG.Code.Url" DefaultValue="https://localhost:3000?commands=1" />
    </bt:Urls>
    <bt:ShortStrings>
        <bt:String id="PG.CodeCommand.Title" DefaultValue="Code" />
    </bt:ShortStrings>
 ... Other markup omitted ...
</Resources>
```

## <a name="supportspinning"></a>SupportsPinning

Elemento opcional quando **xsi:type** for "ShowTaskpane". Os elementos [VersionOverrides](versionoverrides.md) incluídos devem ter um valor `VersionOverridesV1_1` para o atributo `xsi:type`. Inclua esse elemento com um valor `true` a fim de fornecer suporte para fixação do painel de tarefas. O usuário pode "fixar" o painel de tarefas, fazendo com que ele permaneça aberto quando alterar a seleção. Para saber mais, consulte [Implementar um painel de tarefas fixável no Outlook](../../outlook/pinnable-taskpane.md).

> [!IMPORTANT]
> Embora o `SupportsPinning` elemento tenha sido introduzido no [conjunto de requisitos 1,5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), atualmente só há suporte para assinantes do Office 365 usando o seguinte.
> - Outlook 2016 ou posterior no Windows (compilação 7628,1000 ou posterior)
> - Outlook 2016 ou posterior no Mac (Build 16.13.503 ou posterior)
> - Outlook na Web moderno

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
