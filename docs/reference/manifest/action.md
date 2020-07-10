---
title: Elemento Action no arquivo de manifesto
description: Este elemento Especifica a ação a ser executada quando o usuário seleciona um botão ou controle de menu.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 92c783a15d104aba0adb722ab887391b4511ebed
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094446"
---
# <a name="action-element"></a><span data-ttu-id="a7cec-103">Elemento Action</span><span class="sxs-lookup"><span data-stu-id="a7cec-103">Action element</span></span>

<span data-ttu-id="a7cec-104">Especifica a ação a ser executada quando o usuário seleciona um controle de [botão](control.md#button-control) ou de [menu](control.md#menu-dropdown-button-controls) .</span><span class="sxs-lookup"><span data-stu-id="a7cec-104">Specifies the action to perform when the user selects a  [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control.</span></span>

## <a name="attributes"></a><span data-ttu-id="a7cec-105">Atributos</span><span class="sxs-lookup"><span data-stu-id="a7cec-105">Attributes</span></span>

|  <span data-ttu-id="a7cec-106">Atributo</span><span class="sxs-lookup"><span data-stu-id="a7cec-106">Attribute</span></span>  |  <span data-ttu-id="a7cec-107">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="a7cec-107">Required</span></span>  |  <span data-ttu-id="a7cec-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="a7cec-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="a7cec-109">xsi:type</span><span class="sxs-lookup"><span data-stu-id="a7cec-109">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="a7cec-110">Sim</span><span class="sxs-lookup"><span data-stu-id="a7cec-110">Yes</span></span>  | <span data-ttu-id="a7cec-111">Tipo de ação a executar</span><span class="sxs-lookup"><span data-stu-id="a7cec-111">Action type to take</span></span>|

## <a name="child-elements"></a><span data-ttu-id="a7cec-112">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="a7cec-112">Child elements</span></span>

|  <span data-ttu-id="a7cec-113">Elemento</span><span class="sxs-lookup"><span data-stu-id="a7cec-113">Element</span></span> |  <span data-ttu-id="a7cec-114">Descrição</span><span class="sxs-lookup"><span data-stu-id="a7cec-114">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="a7cec-115">FunctionName</span><span class="sxs-lookup"><span data-stu-id="a7cec-115">FunctionName</span></span>](#functionname) |    <span data-ttu-id="a7cec-116">Especifica o nome da função a executar.</span><span class="sxs-lookup"><span data-stu-id="a7cec-116">Specifies the name of the function to execute.</span></span> |
|  [<span data-ttu-id="a7cec-117">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="a7cec-117">SourceLocation</span></span>](#sourcelocation) |    <span data-ttu-id="a7cec-118">Especifica o local do arquivo de origem para essa ação.</span><span class="sxs-lookup"><span data-stu-id="a7cec-118">Specifies the source file location for this action.</span></span> |
| <span data-ttu-id="a7cec-119"> [TaskpaneId](#taskpaneid)</span><span class="sxs-lookup"><span data-stu-id="a7cec-119"> [TaskpaneId](#taskpaneid)</span></span> | <span data-ttu-id="a7cec-120">Especifica a ID do contêiner do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="a7cec-120">Specifies the ID of the task pane container.</span></span>|
| <span data-ttu-id="a7cec-121"> [Title](#title)</span><span class="sxs-lookup"><span data-stu-id="a7cec-121"> [Title](#title)</span></span> | <span data-ttu-id="a7cec-122">Especifica o título personalizado do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="a7cec-122">Specifies the custom title for the task pane.</span></span>|
| <span data-ttu-id="a7cec-123"> [SupportsPinning](#supportspinning)</span><span class="sxs-lookup"><span data-stu-id="a7cec-123"> [SupportsPinning](#supportspinning)</span></span> | <span data-ttu-id="a7cec-124">Especifica se um painel de tarefas tem suporte para fixação, que mantém o painel de tarefas aberto quando o usuário altera a seleção.</span><span class="sxs-lookup"><span data-stu-id="a7cec-124">Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.</span></span>|
  

## <a name="xsitype"></a><span data-ttu-id="a7cec-125">xsi:type</span><span class="sxs-lookup"><span data-stu-id="a7cec-125">xsi:type</span></span>

<span data-ttu-id="a7cec-126">This attribute specifies the kind of action performed when the user selects the button.</span><span class="sxs-lookup"><span data-stu-id="a7cec-126">This attribute specifies the kind of action performed when the user selects the button.</span></span> <span data-ttu-id="a7cec-127">It can be one of the following:</span><span class="sxs-lookup"><span data-stu-id="a7cec-127">It can be one of the following:</span></span>

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a><span data-ttu-id="a7cec-128">FunctionName</span><span class="sxs-lookup"><span data-stu-id="a7cec-128">FunctionName</span></span>

<span data-ttu-id="a7cec-129">Required element when **xsi:type** is "ExecuteFunction".</span><span class="sxs-lookup"><span data-stu-id="a7cec-129">Required element when **xsi:type** is "ExecuteFunction".</span></span> <span data-ttu-id="a7cec-130">Specifies the name of the function to execute.</span><span class="sxs-lookup"><span data-stu-id="a7cec-130">Specifies the name of the function to execute.</span></span> <span data-ttu-id="a7cec-131">The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span><span class="sxs-lookup"><span data-stu-id="a7cec-131">The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span></span>

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a><span data-ttu-id="a7cec-132">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="a7cec-132">SourceLocation</span></span>

<span data-ttu-id="a7cec-133">Elemento obrigatório quando **xsi: Type** for "ShowTaskpane".</span><span class="sxs-lookup"><span data-stu-id="a7cec-133">Required element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="a7cec-134">Especifica o local do arquivo de origem para essa ação.</span><span class="sxs-lookup"><span data-stu-id="a7cec-134">Specifies the source file location for this action.</span></span> <span data-ttu-id="a7cec-135">O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **Url** no elemento **Urls** do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="a7cec-135">The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the [Resources](resources.md) element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a><span data-ttu-id="a7cec-136">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="a7cec-136">TaskpaneId</span></span>

<span data-ttu-id="a7cec-137">Elemento opcional quando  **xsi:type** for "ShowTaskpane".</span><span class="sxs-lookup"><span data-stu-id="a7cec-137">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="a7cec-138">Especifica a ID do contêiner do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="a7cec-138">Specifies the ID of the task pane container.</span></span> <span data-ttu-id="a7cec-139">Quando você tiver várias ações "ShowTaskpane", use uma **TaskpaneId** diferente se desejar ter um painel independente para cada uma.</span><span class="sxs-lookup"><span data-stu-id="a7cec-139">When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each.</span></span> <span data-ttu-id="a7cec-140">Use a mesma **TaskpaneId** para diferentes ações que compartilhem o mesmo painel.</span><span class="sxs-lookup"><span data-stu-id="a7cec-140">Use the same **TaskpaneId** for  different actions that share the same pane.</span></span> <span data-ttu-id="a7cec-141">Quando os usuários escolhem comandos que compartilham a mesma **TaskpaneId**, o contêiner do painel permanece aberto, mas o conteúdo do painel é substituído pela ação "SourceLocation" correspondente.</span><span class="sxs-lookup"><span data-stu-id="a7cec-141">When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation".</span></span>

> [!NOTE]
> <span data-ttu-id="a7cec-142">Esse elemento não tem suporte no Outlook.</span><span class="sxs-lookup"><span data-stu-id="a7cec-142">This element is not supported in Outlook.</span></span>

<span data-ttu-id="a7cec-143">O exemplo a seguir mostra duas ações que compartilham o mesmo **TaskpaneId**.</span><span class="sxs-lookup"><span data-stu-id="a7cec-143">The following example shows two actions that share the same **TaskpaneId**.</span></span>

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

<span data-ttu-id="a7cec-144">The following examples show two actions that use a different **TaskpaneId**.</span><span class="sxs-lookup"><span data-stu-id="a7cec-144">The following examples show two actions that use a different **TaskpaneId**.</span></span> <span data-ttu-id="a7cec-145">To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span><span class="sxs-lookup"><span data-stu-id="a7cec-145">To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span></span>

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

## <a name="title"></a><span data-ttu-id="a7cec-146">Cargo</span><span class="sxs-lookup"><span data-stu-id="a7cec-146">Title</span></span>

<span data-ttu-id="a7cec-147">Elemento opcional quando  **xsi:type** for "ShowTaskpane".</span><span class="sxs-lookup"><span data-stu-id="a7cec-147">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="a7cec-148">Especifica o título personalizado do painel de tarefas desta ação.</span><span class="sxs-lookup"><span data-stu-id="a7cec-148">Specifies the custom title for the task pane for this action.</span></span>

<span data-ttu-id="a7cec-149">O exemplo a seguir mostra uma ação que usa o elemento **title** .</span><span class="sxs-lookup"><span data-stu-id="a7cec-149">The following example shows an action that uses the **Title** element.</span></span> <span data-ttu-id="a7cec-150">Observe que você não atribui o **título** a uma cadeia de caracteres diretamente.</span><span class="sxs-lookup"><span data-stu-id="a7cec-150">Note that you don't assign the **Title** to a string directly.</span></span> <span data-ttu-id="a7cec-151">Em vez disso, atribua um ID de recurso (Resid), que é definido na seção **recursos** do manifesto.</span><span class="sxs-lookup"><span data-stu-id="a7cec-151">Instead, you assign it a resource ID (resid), that is defined in the **Resources** section of the manifest.</span></span>

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

## <a name="supportspinning"></a><span data-ttu-id="a7cec-152">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="a7cec-152">SupportsPinning</span></span>

<span data-ttu-id="a7cec-153">Elemento opcional quando **xsi:type** for "ShowTaskpane".</span><span class="sxs-lookup"><span data-stu-id="a7cec-153">Optional element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="a7cec-154">Os elementos [VersionOverrides](versionoverrides.md) incluídos devem ter um valor `VersionOverridesV1_1` para o atributo `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="a7cec-154">The containing [VersionOverrides](versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span> <span data-ttu-id="a7cec-155">Inclua esse elemento com um valor `true` a fim de fornecer suporte para fixação do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="a7cec-155">Include this element with a value of `true` to support task pane pinning.</span></span> <span data-ttu-id="a7cec-156">O usuário pode "fixar" o painel de tarefas, fazendo com que ele permaneça aberto quando alterar a seleção.</span><span class="sxs-lookup"><span data-stu-id="a7cec-156">The user will be able to "pin" the task pane, causing it to stay open when changing the selection.</span></span> <span data-ttu-id="a7cec-157">Para saber mais, consulte [Implementar um painel de tarefas fixável no Outlook](../../outlook/pinnable-taskpane.md).</span><span class="sxs-lookup"><span data-stu-id="a7cec-157">For more information, see [Implement a pinnable task pane in Outlook](../../outlook/pinnable-taskpane.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a7cec-158">Embora o `SupportsPinning` elemento tenha sido introduzido no [conjunto de requisitos 1,5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), atualmente só há suporte para assinantes do Microsoft 365 usando o seguinte.</span><span class="sxs-lookup"><span data-stu-id="a7cec-158">Although the `SupportsPinning` element was introduced in [requirement set 1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), it's currently only supported for Microsoft 365 subscribers using the following.</span></span>
> - <span data-ttu-id="a7cec-159">Outlook 2016 ou posterior no Windows (compilação 7628,1000 ou posterior)</span><span class="sxs-lookup"><span data-stu-id="a7cec-159">Outlook 2016 or later on Windows (build 7628.1000 or later)</span></span>
> - <span data-ttu-id="a7cec-160">Outlook 2016 ou posterior no Mac (Build 16.13.503 ou posterior)</span><span class="sxs-lookup"><span data-stu-id="a7cec-160">Outlook 2016 or later on Mac (build 16.13.503 or later)</span></span>
> - <span data-ttu-id="a7cec-161">Outlook na Web moderno</span><span class="sxs-lookup"><span data-stu-id="a7cec-161">Modern Outlook on the web</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
