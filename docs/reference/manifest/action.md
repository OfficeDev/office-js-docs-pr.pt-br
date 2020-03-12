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
# <a name="action-element"></a><span data-ttu-id="0df46-102">Elemento Action</span><span class="sxs-lookup"><span data-stu-id="0df46-102">Action element</span></span>

<span data-ttu-id="0df46-103">Especifica a ação a ser executada quando o usuário seleciona controles de [Button](control.md#button-control) ou [Menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="0df46-103">Specifies the action to perform when the user selects a  [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) controls.</span></span>

## <a name="attributes"></a><span data-ttu-id="0df46-104">Atributos</span><span class="sxs-lookup"><span data-stu-id="0df46-104">Attributes</span></span>

|  <span data-ttu-id="0df46-105">Atributo</span><span class="sxs-lookup"><span data-stu-id="0df46-105">Attribute</span></span>  |  <span data-ttu-id="0df46-106">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="0df46-106">Required</span></span>  |  <span data-ttu-id="0df46-107">Descrição</span><span class="sxs-lookup"><span data-stu-id="0df46-107">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="0df46-108">xsi:type</span><span class="sxs-lookup"><span data-stu-id="0df46-108">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="0df46-109">Sim</span><span class="sxs-lookup"><span data-stu-id="0df46-109">Yes</span></span>  | <span data-ttu-id="0df46-110">Tipo de ação a executar</span><span class="sxs-lookup"><span data-stu-id="0df46-110">Action type to take</span></span>|

## <a name="child-elements"></a><span data-ttu-id="0df46-111">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="0df46-111">Child elements</span></span>

|  <span data-ttu-id="0df46-112">Elemento</span><span class="sxs-lookup"><span data-stu-id="0df46-112">Element</span></span> |  <span data-ttu-id="0df46-113">Descrição</span><span class="sxs-lookup"><span data-stu-id="0df46-113">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="0df46-114">FunctionName</span><span class="sxs-lookup"><span data-stu-id="0df46-114">FunctionName</span></span>](#functionname) |    <span data-ttu-id="0df46-115">Especifica o nome da função a executar.</span><span class="sxs-lookup"><span data-stu-id="0df46-115">Specifies the name of the function to execute.</span></span> |
|  [<span data-ttu-id="0df46-116">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="0df46-116">SourceLocation</span></span>](#sourcelocation) |    <span data-ttu-id="0df46-117">Especifica o local do arquivo de origem para essa ação.</span><span class="sxs-lookup"><span data-stu-id="0df46-117">Specifies the source file location for this action.</span></span> |
| <span data-ttu-id="0df46-118"> [TaskpaneId](#taskpaneid)</span><span class="sxs-lookup"><span data-stu-id="0df46-118"> [TaskpaneId](#taskpaneid)</span></span> | <span data-ttu-id="0df46-119">Especifica a ID do contêiner do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="0df46-119">Specifies the ID of the task pane container.</span></span>|
| <span data-ttu-id="0df46-120"> [Title](#title)</span><span class="sxs-lookup"><span data-stu-id="0df46-120"> [Title](#title)</span></span> | <span data-ttu-id="0df46-121">Especifica o título personalizado do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="0df46-121">Specifies the custom title for the task pane.</span></span>|
| <span data-ttu-id="0df46-122"> [SupportsPinning](#supportspinning)</span><span class="sxs-lookup"><span data-stu-id="0df46-122"> [SupportsPinning](#supportspinning)</span></span> | <span data-ttu-id="0df46-123">Especifica se um painel de tarefas tem suporte para fixação, que mantém o painel de tarefas aberto quando o usuário altera a seleção.</span><span class="sxs-lookup"><span data-stu-id="0df46-123">Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.</span></span>|
  

## <a name="xsitype"></a><span data-ttu-id="0df46-124">xsi:type</span><span class="sxs-lookup"><span data-stu-id="0df46-124">xsi:type</span></span>

<span data-ttu-id="0df46-p101">Este atributo especifica o tipo de ação realizada quando o usuário seleciona o botão. Pode ser uma das seguintes:</span><span class="sxs-lookup"><span data-stu-id="0df46-p101">This attribute specifies the kind of action performed when the user selects the button. It can be one of the following:</span></span>

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a><span data-ttu-id="0df46-127">FunctionName</span><span class="sxs-lookup"><span data-stu-id="0df46-127">FunctionName</span></span>

<span data-ttu-id="0df46-p102">Elemento obrigatório quando **xsi:type** é "ExecuteFunction". Especifica o nome da função a ser executada. A função está contida no arquivo especificado no elemento [FunctionFile](functionfile.md).</span><span class="sxs-lookup"><span data-stu-id="0df46-p102">Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span></span>

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a><span data-ttu-id="0df46-131">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="0df46-131">SourceLocation</span></span>

<span data-ttu-id="0df46-132">Elemento obrigatório quando **xsi: Type** for "ShowTaskpane".</span><span class="sxs-lookup"><span data-stu-id="0df46-132">Required element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="0df46-133">Especifica o local do arquivo de origem para essa ação.</span><span class="sxs-lookup"><span data-stu-id="0df46-133">Specifies the source file location for this action.</span></span> <span data-ttu-id="0df46-134">O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **Url** no elemento **Urls** do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="0df46-134">The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the [Resources](resources.md) element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a><span data-ttu-id="0df46-135">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="0df46-135">TaskpaneId</span></span>

<span data-ttu-id="0df46-136">Elemento opcional quando  **xsi:type** for "ShowTaskpane".</span><span class="sxs-lookup"><span data-stu-id="0df46-136">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="0df46-137">Especifica a ID do contêiner do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="0df46-137">Specifies the ID of the task pane container.</span></span> <span data-ttu-id="0df46-138">Quando você tiver várias ações "ShowTaskpane", use uma **TaskpaneId** diferente se desejar ter um painel independente para cada uma.</span><span class="sxs-lookup"><span data-stu-id="0df46-138">When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each.</span></span> <span data-ttu-id="0df46-139">Use a mesma **TaskpaneId** para diferentes ações que compartilhem o mesmo painel.</span><span class="sxs-lookup"><span data-stu-id="0df46-139">Use the same **TaskpaneId** for  different actions that share the same pane.</span></span> <span data-ttu-id="0df46-140">Quando os usuários escolhem comandos que compartilham a mesma **TaskpaneId**, o contêiner do painel permanece aberto, mas o conteúdo do painel é substituído pela ação "SourceLocation" correspondente.</span><span class="sxs-lookup"><span data-stu-id="0df46-140">When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation".</span></span>

> [!NOTE]
> <span data-ttu-id="0df46-141">Esse elemento não tem suporte no Outlook.</span><span class="sxs-lookup"><span data-stu-id="0df46-141">This element is not supported in Outlook.</span></span>

<span data-ttu-id="0df46-142">O exemplo a seguir mostra duas ações que compartilham o mesmo **TaskpaneId**.</span><span class="sxs-lookup"><span data-stu-id="0df46-142">The following example shows two actions that share the same **TaskpaneId**.</span></span>

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

<span data-ttu-id="0df46-p105">O exemplo a seguir mostra duas ações que usam um **TaskpaneId** diferente. Para ver esses exemplos em contexto, consulte [Exemplo de comando de suplemento simples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span><span class="sxs-lookup"><span data-stu-id="0df46-p105">The following examples show two actions that use a different **TaskpaneId**. To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span></span>

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

## <a name="title"></a><span data-ttu-id="0df46-145">Cargo</span><span class="sxs-lookup"><span data-stu-id="0df46-145">Title</span></span>

<span data-ttu-id="0df46-146">Elemento opcional quando  **xsi:type** for "ShowTaskpane".</span><span class="sxs-lookup"><span data-stu-id="0df46-146">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="0df46-147">Especifica o título personalizado do painel de tarefas desta ação.</span><span class="sxs-lookup"><span data-stu-id="0df46-147">Specifies the custom title for the task pane for this action.</span></span>

<span data-ttu-id="0df46-148">O exemplo a seguir mostra uma ação que usa o elemento **title** .</span><span class="sxs-lookup"><span data-stu-id="0df46-148">The following example shows an action that uses the **Title** element.</span></span> <span data-ttu-id="0df46-149">Observe que você não atribui o **título** a uma cadeia de caracteres diretamente.</span><span class="sxs-lookup"><span data-stu-id="0df46-149">Note that you don't assign the **Title** to a string directly.</span></span> <span data-ttu-id="0df46-150">Em vez disso, atribua um ID de recurso (Resid), que é definido na seção **recursos** do manifesto.</span><span class="sxs-lookup"><span data-stu-id="0df46-150">Instead, you assign it a resource ID (resid), that is defined in the **Resources** section of the manifest.</span></span>

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

## <a name="supportspinning"></a><span data-ttu-id="0df46-151">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="0df46-151">SupportsPinning</span></span>

<span data-ttu-id="0df46-152">Elemento opcional quando **xsi:type** for "ShowTaskpane".</span><span class="sxs-lookup"><span data-stu-id="0df46-152">Optional element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="0df46-153">Os elementos [VersionOverrides](versionoverrides.md) incluídos devem ter um valor `VersionOverridesV1_1` para o atributo `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="0df46-153">The containing [VersionOverrides](versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span> <span data-ttu-id="0df46-154">Inclua esse elemento com um valor `true` a fim de fornecer suporte para fixação do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="0df46-154">Include this element with a value of `true` to support task pane pinning.</span></span> <span data-ttu-id="0df46-155">O usuário pode "fixar" o painel de tarefas, fazendo com que ele permaneça aberto quando alterar a seleção.</span><span class="sxs-lookup"><span data-stu-id="0df46-155">The user will be able to "pin" the task pane, causing it to stay open when changing the selection.</span></span> <span data-ttu-id="0df46-156">Para saber mais, consulte [Implementar um painel de tarefas fixável no Outlook](../../outlook/pinnable-taskpane.md).</span><span class="sxs-lookup"><span data-stu-id="0df46-156">For more information, see [Implement a pinnable task pane in Outlook](../../outlook/pinnable-taskpane.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0df46-157">Embora o `SupportsPinning` elemento tenha sido introduzido no [conjunto de requisitos 1,5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), atualmente só há suporte para assinantes do Office 365 usando o seguinte.</span><span class="sxs-lookup"><span data-stu-id="0df46-157">Although the `SupportsPinning` element was introduced in [requirement set 1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), it's currently only supported for Office 365 subscribers using the following.</span></span>
> - <span data-ttu-id="0df46-158">Outlook 2016 ou posterior no Windows (compilação 7628,1000 ou posterior)</span><span class="sxs-lookup"><span data-stu-id="0df46-158">Outlook 2016 or later on Windows (build 7628.1000 or later)</span></span>
> - <span data-ttu-id="0df46-159">Outlook 2016 ou posterior no Mac (Build 16.13.503 ou posterior)</span><span class="sxs-lookup"><span data-stu-id="0df46-159">Outlook 2016 or later on Mac (build 16.13.503 or later)</span></span>
> - <span data-ttu-id="0df46-160">Outlook na Web moderno</span><span class="sxs-lookup"><span data-stu-id="0df46-160">Modern Outlook on the web</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
