---
title: Elemento Action no arquivo de manifesto
description: Esse elemento especifica a ação a ser executar quando o usuário seleciona um botão ou um controle de menu.
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 21c8f9a6345641f23aad70efed67c9c45f72a1c8
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340411"
---
# <a name="action-element"></a>Elemento Action

Especifica a ação a ser executar quando o usuário seleciona um  [controle Button](control-button.md) ou [Menu](control-menu.md) .

**Válido somente nesses esquemas VersionOverrides**:

- Painel de tarefas 1.0
- Email 1.0
- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) quando o **VersionOverrides** pai é o tipo Taskpane 1.0.
- [Caixa de correio 1.3](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) quando o **VersionOverrides** pai é o tipo Mail 1.0.
- [Caixa de correio 1.5](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) quando o **VersionOverrides** pai é o tipo Mail 1.1.

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Sim  | Tipo de ação a executar|

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Descrição  |
|:-----|:-----|
|  [FunctionName](#functionname) |    Especifica o nome da função a executar. |
|  [SourceLocation](#sourcelocation) |    Especifica o local do arquivo de origem para essa ação. |
|  [TaskpaneId](#taskpaneid) | Especifica a ID do contêiner do painel de tarefas. Não há suporte em Outlook de complementos.|
|  [Title](#title) | Especifica o título personalizado do painel de tarefas. Não há suporte em Outlook de complementos.|
|  [SupportsPinning](#supportspinning) | Especifica se um painel de tarefas tem suporte para fixação, que mantém o painel de tarefas aberto quando o usuário altera a seleção.|

## <a name="xsitype"></a>xsi:type

Este atributo especifica o tipo de ação realizada quando o usuário seleciona o botão. Pode ser uma das seguintes:

- `ExecuteFunction`
- `ShowTaskpane`

> [!IMPORTANT]
> O registro [de eventos de](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) Caixa de Correio e [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) não está disponível quando **xsi:type** é `ExecuteFunction`.

## <a name="functionname"></a>FunctionName

Elemento obrigatório quando **xsi:type** é `ExecuteFunction`. Especifica o nome da função a ser executada. A função está contida no arquivo especificado no elemento [FunctionFile](functionfile.md).

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a>SourceLocation

Elemento obrigatório quando **xsi:type** é `ShowTaskpane`. Especifica o local do arquivo de origem para essa ação. O **atributo resid** não pode ter mais de 32 caracteres e deve ser definido como o valor do atributo **id** de um **elemento Url** no elemento **Urls** no elemento [Resources](resources.md) .

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a>TaskpaneId

Elemento opcional quando  **xsi:type** é `ShowTaskpane`. Especifica a ID do contêiner do painel de tarefas. Quando você tiver várias `ShowTaskpane` ações, use um **TaskpaneId** diferente se quiser um painel independente para cada um. Use a mesma **TaskpaneId** para diferentes ações que compartilhem o mesmo painel. Quando os usuários escolhem comandos que compartilham o mesmo **TaskpaneId**, o contêiner do painel permanecerá aberto, mas o conteúdo do painel será substituído pela Ação correspondente `SourceLocation`.

**Tipo de suplemento:** Painel de tarefas

**Válido somente nesses esquemas VersionOverrides**:

- Painel de tarefas 1.0

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md)

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

Elemento opcional quando  **xsi:type** é `ShowTaskpane`. Especifica o título personalizado do painel de tarefas desta ação.

**Tipo de suplemento:** Painel de tarefas

**Válido somente nesses esquemas VersionOverrides**:

- Painel de tarefas 1.0

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md)

> [!NOTE]
> Esse elemento filho não é suportado em Outlook de complementos.

O exemplo a seguir mostra uma ação que usa o **elemento Title** . Observe que você não atribui o **Título** a uma cadeia de caracteres diretamente. Em vez disso, você atribui a ele uma ID de recurso (resid), que é definida na seção **Recursos** do manifesto e não pode ter mais de 32 caracteres.

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

Elemento opcional quando **xsi:type** é `ShowTaskpane`. Os elementos [VersionOverrides](versionoverrides.md) que contêm devem ter um valor de atributo **xsi:type** de `VersionOverridesV1_1`. Inclua esse elemento com um valor `true` a fim de fornecer suporte para fixação do painel de tarefas. O usuário pode "fixar" o painel de tarefas, fazendo com que ele permaneça aberto quando alterar a seleção. Para saber mais, consulte [Implementar um painel de tarefas fixável no Outlook](../../outlook/pinnable-taskpane.md).

**Tipo de suplemento:** Email

**Válido somente nesses esquemas VersionOverrides**:

- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [Caixa de correio 1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)

> [!IMPORTANT]
> Embora o **elemento SupportsPinning** tenha sido introduzido no conjunto de requisitos [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), ele atualmente só tem suporte para assinantes Microsoft 365 usando o seguinte:
>
> - Outlook 2016 ou posterior no Windows (build 7628.1000 ou posterior)
> - Outlook 2016 ou posterior no Mac (build 16.13.503 ou posterior)
> - Outlook na Web moderno

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
