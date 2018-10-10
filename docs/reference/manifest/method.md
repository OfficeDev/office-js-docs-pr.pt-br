# <a name="method-element"></a><span data-ttu-id="ab62d-101">Elemento Method</span><span class="sxs-lookup"><span data-stu-id="ab62d-101">Method element</span></span>

<span data-ttu-id="ab62d-102">Especifica um método individual a partir da API do JavaScript para Office que o Suplemento do Office exige para ativar.</span><span class="sxs-lookup"><span data-stu-id="ab62d-102">Specifies an individual method from the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="ab62d-103">**Tipo de suplemento:** Conteúdo, Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="ab62d-103">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="ab62d-104">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="ab62d-104">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="ab62d-105">Contido em</span><span class="sxs-lookup"><span data-stu-id="ab62d-105">Contained in:</span></span>

[<span data-ttu-id="ab62d-106">Métodos</span><span class="sxs-lookup"><span data-stu-id="ab62d-106">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="ab62d-107">Atributos</span><span class="sxs-lookup"><span data-stu-id="ab62d-107">Attributes</span></span>

|<span data-ttu-id="ab62d-108">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="ab62d-108">**Attribute**</span></span>|<span data-ttu-id="ab62d-109">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="ab62d-109">**Type**</span></span>|<span data-ttu-id="ab62d-110">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="ab62d-110">**Required**</span></span>|<span data-ttu-id="ab62d-111">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="ab62d-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="ab62d-112">Nome</span><span class="sxs-lookup"><span data-stu-id="ab62d-112">Name</span></span>|<span data-ttu-id="ab62d-113">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ab62d-113">string</span></span>|<span data-ttu-id="ab62d-114">obrigatório</span><span class="sxs-lookup"><span data-stu-id="ab62d-114">required</span></span>|<span data-ttu-id="ab62d-p101">Especifica o nome do método necessário qualificado com seu objeto pai. Por exemplo, para especificar o método **getSelectedDataAsync**, você deve especificar `"Document.getSelectedDataAsync"`.</span><span class="sxs-lookup"><span data-stu-id="ab62d-p101">Specifies the name of the required method qualified with its parent object. For example, to specify the  **getSelectedDataAsync** method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="ab62d-117">Comentários</span><span class="sxs-lookup"><span data-stu-id="ab62d-117">Remarks</span></span>

<span data-ttu-id="ab62d-118">Os elementos **Methods** e **Method** não são compatíveis com os suplementos de email. Para obter mais informações sobre os conjuntos de requisitos, consulte [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="ab62d-118">The  Methods and Method elements aren't supported by mail add-ins. For more information about requirement sets, see Specify Office hosts and API requirements.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="ab62d-119">Como não há forma de especificar o requisito de versão mínima de métodos individuais, para verificar se um método está disponível em tempo de execução, você também deve usar uma instrução **if** ao chamar o método no script do suplemento.</span><span class="sxs-lookup"><span data-stu-id="ab62d-119">Important  Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an  **if** statement when calling that method in the script of your add-in. For more information about how to do this, see Understanding the JavaScript API for Office.</span></span> <span data-ttu-id="ab62d-120">Para obter mais informações sobre como fazer isso, consulte [Noções básicas sobre a API do JavaScript para Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span><span class="sxs-lookup"><span data-stu-id="ab62d-120">For more information about how to do this, see [Understanding the JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span></span>

