# <a name="action-element"></a><span data-ttu-id="c0bae-101">Elemento Action</span><span class="sxs-lookup"><span data-stu-id="c0bae-101">Action element</span></span>

<span data-ttu-id="c0bae-102">Especifica a ação a ser executada quando o usuário seleciona um controle de [Botão](control.md#button-control) ou [Menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="c0bae-102">Specifies the action to perform when the user selects a  [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) controls.</span></span>
 
## <a name="attributes"></a><span data-ttu-id="c0bae-103">Atributos</span><span class="sxs-lookup"><span data-stu-id="c0bae-103">Attributes</span></span>

|  <span data-ttu-id="c0bae-104">Atributo</span><span class="sxs-lookup"><span data-stu-id="c0bae-104">Attribute</span></span>  |  <span data-ttu-id="c0bae-105">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="c0bae-105">Required</span></span>  |  <span data-ttu-id="c0bae-106">Descrição</span><span class="sxs-lookup"><span data-stu-id="c0bae-106">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="c0bae-107">xsi:type</span><span class="sxs-lookup"><span data-stu-id="c0bae-107">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="c0bae-108">Sim</span><span class="sxs-lookup"><span data-stu-id="c0bae-108">Yes</span></span>  | <span data-ttu-id="c0bae-109">Tipo de ação que será executada</span><span class="sxs-lookup"><span data-stu-id="c0bae-109">Action type to take</span></span>|

## <a name="child-elements"></a><span data-ttu-id="c0bae-110">Elementos filhos</span><span class="sxs-lookup"><span data-stu-id="c0bae-110">Child elements</span></span>

|  <span data-ttu-id="c0bae-111">Elemento</span><span class="sxs-lookup"><span data-stu-id="c0bae-111">Element</span></span> |  <span data-ttu-id="c0bae-112">Descrição</span><span class="sxs-lookup"><span data-stu-id="c0bae-112">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="c0bae-113">FunctionName</span><span class="sxs-lookup"><span data-stu-id="c0bae-113">FunctionName</span></span>](#functionname) |    <span data-ttu-id="c0bae-114">Especifica o nome da função que será executada.</span><span class="sxs-lookup"><span data-stu-id="c0bae-114">Specifies the name of the function to execute.</span></span> |
|  [<span data-ttu-id="c0bae-115">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="c0bae-115">SourceLocation</span></span>](#sourcelocation) |    <span data-ttu-id="c0bae-116">Especifica o local do arquivo de origem para essa ação.</span><span class="sxs-lookup"><span data-stu-id="c0bae-116">Specifies the source file location for this action.</span></span> |
|  [<span data-ttu-id="c0bae-117">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="c0bae-117">TaskpaneId</span></span>](#taskpaneid) | <span data-ttu-id="c0bae-118">Especifica a ID do contêiner do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="c0bae-118">Specifies the ID of the task pane container.</span></span>|
|  [<span data-ttu-id="c0bae-119">Title</span><span class="sxs-lookup"><span data-stu-id="c0bae-119">Title</span></span>](#title) | <span data-ttu-id="c0bae-120">Especifica o título personalizado do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="c0bae-120">Specifies the custom title for the task pane.</span></span>|
|  [<span data-ttu-id="c0bae-121">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="c0bae-121">SupportsPinning</span></span>](#supportspinning) | <span data-ttu-id="c0bae-122">Especifica se um painel de tarefas tem suporte para fixação, o que mantém o painel de tarefas aberto quando o usuário altera a seleção.</span><span class="sxs-lookup"><span data-stu-id="c0bae-122">Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.</span></span>|
  

## <a name="xsitype"></a><span data-ttu-id="c0bae-123">xsi:type</span><span class="sxs-lookup"><span data-stu-id="c0bae-123">xsi:type</span></span>

<span data-ttu-id="c0bae-p101">Este atributo especifica o tipo de ação realizada quando o usuário seleciona o botão. Pode ser uma das seguintes:</span><span class="sxs-lookup"><span data-stu-id="c0bae-p101">This attribute specifies the kind of action performed when the user selects the button. It can be one of the following:</span></span>

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a><span data-ttu-id="c0bae-126">FunctionName</span><span class="sxs-lookup"><span data-stu-id="c0bae-126">FunctionName</span></span>

<span data-ttu-id="c0bae-p102">Elemento obrigatório quando **xsi:type** for "ExecuteFunction". Especifica o nome da função que será executada. A função está contida no arquivo especificado no elemento [FunctionFile](functionfile.md).</span><span class="sxs-lookup"><span data-stu-id="c0bae-p102">Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span></span>

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a><span data-ttu-id="c0bae-130">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="c0bae-130">SourceLocation</span></span>

<span data-ttu-id="c0bae-p103">Elemento obrigatório quando **xsi:type** for "ShowTaskpane". Especifica o local do arquivo de origem para essa ação. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **Url** no elemento **Urls** no elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="c0bae-p103">Required element when  **xsi:type** is "ShowTaskpane". Specifies the source file location for this action. The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the [Resources](resources.md) element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a><span data-ttu-id="c0bae-134">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="c0bae-134">TaskpaneId</span></span>

<span data-ttu-id="c0bae-p104">Elemento opcional quando **xsi:type** for "ShowTaskpane". Especifica a ID do contêiner do painel de tarefas. Quando você tiver várias ações "ShowTaskpane", use uma **TaskpaneId** diferente se desejar ter um painel independente para cada uma. Use a mesma **TaskpaneId** para diferentes ações que compartilham o mesmo painel. Quando os usuários escolhem comandos que compartilham a mesma **TaskpaneId**, o contêiner do painel permanece aberto, porém o conteúdo do painel é substituído pela Ação "SourceLocation" correspondente.</span><span class="sxs-lookup"><span data-stu-id="c0bae-p104">Optional element when  **xsi:type** is "ShowTaskpane". Specifies the ID of the task pane container. When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each. Use the same **TaskpaneId** for  different actions that share the same pane. When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation".</span></span> 

> [!NOTE]
> <span data-ttu-id="c0bae-140">Esse elemento não tem suporte no Outlook.</span><span class="sxs-lookup"><span data-stu-id="c0bae-140">Note: This element is not supported in Outlook.</span></span>

<span data-ttu-id="c0bae-141">O exemplo a seguir mostra duas ações que compartilham a mesma **TaskpaneId**.</span><span class="sxs-lookup"><span data-stu-id="c0bae-141">The following example shows two actions that share the same **TaskpaneId**.</span></span> 

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

<span data-ttu-id="c0bae-p105">O exemplo a seguir mostra duas ações que usam uma **TaskpaneId** diferente. Para ver esses exemplos em contexto, confira [Exemplo de comandos de suplemento simples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span><span class="sxs-lookup"><span data-stu-id="c0bae-p105">The following examples show two actions that use a different **TaskpaneId**. To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span></span>

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

## <a name="title"></a><span data-ttu-id="c0bae-144">Title</span><span class="sxs-lookup"><span data-stu-id="c0bae-144">Title</span></span>
<span data-ttu-id="c0bae-p106">Elemento opcional quando **xsi:type**  for "ShowTaskpane". Especifica o título personalizado do painel de tarefas desta ação.</span><span class="sxs-lookup"><span data-stu-id="c0bae-p106">Optional element when  **xsi:type** is "ShowTaskpane". Specifies the custom title for the task pane for this action.</span></span> 

<span data-ttu-id="c0bae-147">Os exemplos a seguir mostram duas ações distintas que usam o elemento **Title**.</span><span class="sxs-lookup"><span data-stu-id="c0bae-147">The following examples show two different actions that use the **Title** element.</span></span>

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

## <a name="supportspinning"></a><span data-ttu-id="c0bae-148">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="c0bae-148">SupportsPinning</span></span>

<span data-ttu-id="c0bae-p107">Elemento opcional quando **xsi:type** for "ShowTaskpane". Os elementos que contêm [VersionOverrides](versionoverrides.md) devem ter um valor de atributo `xsi:type` de `VersionOverridesV1_1`. Inclua esse elemento com um valor de `true` a fim de fornecer suporte para fixação do painel de tarefas. O usuário poderá "fixar" o painel de tarefas, fazendo com que ele permaneça aberto quando alterar a seleção. Para saber mais, confira [Implementar um painel de tarefas fixável no Outlook](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane).</span><span class="sxs-lookup"><span data-stu-id="c0bae-p107">Optional element when **xsi:type** is "ShowTaskpane". The containing [VersionOverrides](versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`. Include this element with a value of `true` to support taskpane pinning. The user will be able to "pin" the taskpane, causing it to stay open when changing the selection. For more information, see [Implement a pinnable taskpane in Outlook](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane).</span></span>

> [!NOTE]
> <span data-ttu-id="c0bae-154">Atualmente, o SupportsPinning só é suportado pelo Outlook 2016 para Windows (build 7628.1000 ou posterior).</span><span class="sxs-lookup"><span data-stu-id="c0bae-154">Note: SupportsPinning currently only supported by Outlook 2016 for Windows (build 7628.1000 or later).</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```


