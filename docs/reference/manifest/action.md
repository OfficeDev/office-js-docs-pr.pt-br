---
title: Elemento Action no arquivo de manifesto
description: Esse elemento especifica a ação a ser executar quando o usuário seleciona um botão ou um controle de menu.
ms.date: 02/12/2021
localization_priority: Normal
ms.openlocfilehash: 6be1430800dea27dbd9bf78607161d88e475c145
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505406"
---
# <a name="action-element"></a><span data-ttu-id="dc305-103">Elemento Action</span><span class="sxs-lookup"><span data-stu-id="dc305-103">Action element</span></span>

<span data-ttu-id="dc305-104">Especifica a ação a ser executar quando o usuário seleciona um [controle Button](control.md#button-control) ou [Menu.](control.md#menu-dropdown-button-controls)</span><span class="sxs-lookup"><span data-stu-id="dc305-104">Specifies the action to perform when the user selects a  [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control.</span></span>

## <a name="attributes"></a><span data-ttu-id="dc305-105">Atributos</span><span class="sxs-lookup"><span data-stu-id="dc305-105">Attributes</span></span>

|  <span data-ttu-id="dc305-106">Atributo</span><span class="sxs-lookup"><span data-stu-id="dc305-106">Attribute</span></span>  |  <span data-ttu-id="dc305-107">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="dc305-107">Required</span></span>  |  <span data-ttu-id="dc305-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="dc305-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="dc305-109">xsi:type</span><span class="sxs-lookup"><span data-stu-id="dc305-109">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="dc305-110">Sim</span><span class="sxs-lookup"><span data-stu-id="dc305-110">Yes</span></span>  | <span data-ttu-id="dc305-111">Tipo de ação a executar</span><span class="sxs-lookup"><span data-stu-id="dc305-111">Action type to take</span></span>|

## <a name="child-elements"></a><span data-ttu-id="dc305-112">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="dc305-112">Child elements</span></span>

|  <span data-ttu-id="dc305-113">Elemento</span><span class="sxs-lookup"><span data-stu-id="dc305-113">Element</span></span> |  <span data-ttu-id="dc305-114">Descrição</span><span class="sxs-lookup"><span data-stu-id="dc305-114">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="dc305-115">FunctionName</span><span class="sxs-lookup"><span data-stu-id="dc305-115">FunctionName</span></span>](#functionname) |    <span data-ttu-id="dc305-116">Especifica o nome da função a executar.</span><span class="sxs-lookup"><span data-stu-id="dc305-116">Specifies the name of the function to execute.</span></span> |
|  [<span data-ttu-id="dc305-117">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="dc305-117">SourceLocation</span></span>](#sourcelocation) |    <span data-ttu-id="dc305-118">Especifica o local do arquivo de origem para essa ação.</span><span class="sxs-lookup"><span data-stu-id="dc305-118">Specifies the source file location for this action.</span></span> |
|  [<span data-ttu-id="dc305-119">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="dc305-119">TaskpaneId</span></span>](#taskpaneid) | <span data-ttu-id="dc305-120">Especifica a ID do contêiner do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="dc305-120">Specifies the ID of the task pane container.</span></span>|
|  [<span data-ttu-id="dc305-121">Title</span><span class="sxs-lookup"><span data-stu-id="dc305-121">Title</span></span>](#title) | <span data-ttu-id="dc305-122">Especifica o título personalizado do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="dc305-122">Specifies the custom title for the task pane.</span></span>|
|  [<span data-ttu-id="dc305-123">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="dc305-123">SupportsPinning</span></span>](#supportspinning) | <span data-ttu-id="dc305-124">Especifica se um painel de tarefas tem suporte para fixação, que mantém o painel de tarefas aberto quando o usuário altera a seleção.</span><span class="sxs-lookup"><span data-stu-id="dc305-124">Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.</span></span>|

## <a name="xsitype"></a><span data-ttu-id="dc305-125">xsi:type</span><span class="sxs-lookup"><span data-stu-id="dc305-125">xsi:type</span></span>

<span data-ttu-id="dc305-p101">Este atributo especifica o tipo de ação realizada quando o usuário seleciona o botão. Pode ser uma das seguintes:</span><span class="sxs-lookup"><span data-stu-id="dc305-p101">This attribute specifies the kind of action performed when the user selects the button. It can be one of the following:</span></span>

- `ExecuteFunction`
- `ShowTaskpane`

> [!IMPORTANT]
> <span data-ttu-id="dc305-128">O registro [de eventos de](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) Caixa de Correio e [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) não está disponível quando **xsi:type** é `ExecuteFunction` .</span><span class="sxs-lookup"><span data-stu-id="dc305-128">Registering [Mailbox](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) and [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) events is not available when **xsi:type** is `ExecuteFunction`.</span></span>

## <a name="functionname"></a><span data-ttu-id="dc305-129">FunctionName</span><span class="sxs-lookup"><span data-stu-id="dc305-129">FunctionName</span></span>

<span data-ttu-id="dc305-p102">Elemento obrigatório quando **xsi:type** é "ExecuteFunction". Especifica o nome da função a ser executada. A função está contida no arquivo especificado no elemento [FunctionFile](functionfile.md).</span><span class="sxs-lookup"><span data-stu-id="dc305-p102">Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span></span>

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a><span data-ttu-id="dc305-133">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="dc305-133">SourceLocation</span></span>

<span data-ttu-id="dc305-134">Elemento obrigatório quando **xsi:type** é "ShowTaskpane".</span><span class="sxs-lookup"><span data-stu-id="dc305-134">Required element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="dc305-135">Especifica o local do arquivo de origem para essa ação.</span><span class="sxs-lookup"><span data-stu-id="dc305-135">Specifies the source file location for this action.</span></span> <span data-ttu-id="dc305-136">O **atributo resid** não pode ter mais de 32 caracteres e deve ser definido como o valor do atributo **id** de um elemento Url no elemento **Urls** no elemento [Resources.](resources.md) </span><span class="sxs-lookup"><span data-stu-id="dc305-136">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the [Resources](resources.md) element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a><span data-ttu-id="dc305-137">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="dc305-137">TaskpaneId</span></span>

<span data-ttu-id="dc305-p104">Elemento opcional quando **xsi:type** for "ShowTaskpane". Especifica a ID do contêiner do painel de tarefas. Quando você tiver várias ações "ShowTaskpane", use uma **TaskpaneId** diferente se desejar ter um painel independente para cada uma. Use a mesma **TaskpaneId** para diferentes ações que compartilhem o mesmo painel. Quando os usuários escolhem comandos que compartilham o mesmo **TaskpaneId**, o contêiner do painel permanece aberto, mas o conteúdo do painel é substituído pela ação correspondente "SourceLocation".</span><span class="sxs-lookup"><span data-stu-id="dc305-p104">Optional element when  **xsi:type** is "ShowTaskpane". Specifies the ID of the task pane container. When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each. Use the same **TaskpaneId** for  different actions that share the same pane. When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation".</span></span>

> [!NOTE]
> <span data-ttu-id="dc305-143">Esse elemento não tem suporte no Outlook.</span><span class="sxs-lookup"><span data-stu-id="dc305-143">This element is not supported in Outlook.</span></span>

<span data-ttu-id="dc305-144">O exemplo a seguir mostra duas ações que compartilham o mesmo **TaskpaneId**.</span><span class="sxs-lookup"><span data-stu-id="dc305-144">The following example shows two actions that share the same **TaskpaneId**.</span></span>

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

<span data-ttu-id="dc305-p105">O exemplo a seguir mostra duas ações que usam um **TaskpaneId** diferente. Para ver esses exemplos em contexto, consulte [Exemplo de comando de suplemento simples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span><span class="sxs-lookup"><span data-stu-id="dc305-p105">The following examples show two actions that use a different **TaskpaneId**. To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span></span>

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

## <a name="title"></a><span data-ttu-id="dc305-147">Cargo</span><span class="sxs-lookup"><span data-stu-id="dc305-147">Title</span></span>

<span data-ttu-id="dc305-148">Elemento opcional quando **xsi:type** for "ShowTaskpane".</span><span class="sxs-lookup"><span data-stu-id="dc305-148">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="dc305-149">Especifica o título personalizado do painel de tarefas desta ação.</span><span class="sxs-lookup"><span data-stu-id="dc305-149">Specifies the custom title for the task pane for this action.</span></span>

<span data-ttu-id="dc305-150">O exemplo a seguir mostra uma ação que usa o **elemento Title.**</span><span class="sxs-lookup"><span data-stu-id="dc305-150">The following example shows an action that uses the **Title** element.</span></span> <span data-ttu-id="dc305-151">Observe que você não atribui o **Título** a uma cadeia de caracteres diretamente.</span><span class="sxs-lookup"><span data-stu-id="dc305-151">Note that you don't assign the **Title** to a string directly.</span></span> <span data-ttu-id="dc305-152">Em vez disso, você atribui a ele uma ID de recurso (resid), que é definida na seção **Recursos** do manifesto e não pode ter mais de 32 caracteres.</span><span class="sxs-lookup"><span data-stu-id="dc305-152">Instead, you assign it a resource ID (resid), that is defined in the **Resources** section of the manifest and can be no more than 32 characters.</span></span>

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

## <a name="supportspinning"></a><span data-ttu-id="dc305-153">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="dc305-153">SupportsPinning</span></span>

<span data-ttu-id="dc305-154">Elemento opcional quando **xsi:type** for "ShowTaskpane".</span><span class="sxs-lookup"><span data-stu-id="dc305-154">Optional element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="dc305-155">Os elementos [VersionOverrides](versionoverrides.md) incluídos devem ter um valor `VersionOverridesV1_1` para o atributo `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="dc305-155">The containing [VersionOverrides](versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span> <span data-ttu-id="dc305-156">Inclua esse elemento com um valor `true` a fim de fornecer suporte para fixação do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="dc305-156">Include this element with a value of `true` to support task pane pinning.</span></span> <span data-ttu-id="dc305-157">O usuário pode "fixar" o painel de tarefas, fazendo com que ele permaneça aberto quando alterar a seleção.</span><span class="sxs-lookup"><span data-stu-id="dc305-157">The user will be able to "pin" the task pane, causing it to stay open when changing the selection.</span></span> <span data-ttu-id="dc305-158">Para saber mais, consulte [Implementar um painel de tarefas fixável no Outlook](../../outlook/pinnable-taskpane.md).</span><span class="sxs-lookup"><span data-stu-id="dc305-158">For more information, see [Implement a pinnable task pane in Outlook](../../outlook/pinnable-taskpane.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="dc305-159">Embora o elemento tenha sido introduzido no conjunto de requisitos `SupportsPinning` [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), ele atualmente só tem suporte para assinantes do Microsoft 365 usando o seguinte.</span><span class="sxs-lookup"><span data-stu-id="dc305-159">Although the `SupportsPinning` element was introduced in [requirement set 1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), it's currently only supported for Microsoft 365 subscribers using the following.</span></span>
>
> - <span data-ttu-id="dc305-160">Outlook 2016 ou posterior no Windows (build 7628.1000 ou posterior)</span><span class="sxs-lookup"><span data-stu-id="dc305-160">Outlook 2016 or later on Windows (build 7628.1000 or later)</span></span>
> - <span data-ttu-id="dc305-161">Outlook 2016 ou posterior no Mac (build 16.13.503 ou posterior)</span><span class="sxs-lookup"><span data-stu-id="dc305-161">Outlook 2016 or later on Mac (build 16.13.503 or later)</span></span>
> - <span data-ttu-id="dc305-162">Outlook na Web moderno</span><span class="sxs-lookup"><span data-stu-id="dc305-162">Modern Outlook on the web</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
