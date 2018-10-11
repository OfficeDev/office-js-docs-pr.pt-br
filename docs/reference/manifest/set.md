# <a name="set-element"></a><span data-ttu-id="6461c-101">Elemento Set</span><span class="sxs-lookup"><span data-stu-id="6461c-101">Set element</span></span>

<span data-ttu-id="6461c-102">Especifica um conjunto de requisitos a partir da API JavaScript para Office que o seu suplemento do Office exige para ativar.</span><span class="sxs-lookup"><span data-stu-id="6461c-102">Specifies a requirement set from the JavaScript API for Office that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="6461c-103">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="6461c-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="6461c-104">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="6461c-104">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="6461c-105">Contido em</span><span class="sxs-lookup"><span data-stu-id="6461c-105">Contained in:</span></span>

[<span data-ttu-id="6461c-106">Conjuntos</span><span class="sxs-lookup"><span data-stu-id="6461c-106">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="6461c-107">Atributos</span><span class="sxs-lookup"><span data-stu-id="6461c-107">Attributes</span></span>

|<span data-ttu-id="6461c-108">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="6461c-108">**Attribute**</span></span>|<span data-ttu-id="6461c-109">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="6461c-109">**Type**</span></span>|<span data-ttu-id="6461c-110">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="6461c-110">**Required**</span></span>|<span data-ttu-id="6461c-111">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="6461c-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="6461c-112">Nome</span><span class="sxs-lookup"><span data-stu-id="6461c-112">Name</span></span>|<span data-ttu-id="6461c-113">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="6461c-113">string</span></span>|<span data-ttu-id="6461c-114">obrigatório</span><span class="sxs-lookup"><span data-stu-id="6461c-114">required</span></span>|<span data-ttu-id="6461c-115">O nome de um [conjunto de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="6461c-115">The name of a [requirement set](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>|
|<span data-ttu-id="6461c-116">MinVersion</span><span class="sxs-lookup"><span data-stu-id="6461c-116">MinVersion</span></span>|<span data-ttu-id="6461c-117">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="6461c-117">string</span></span>|<span data-ttu-id="6461c-118">opcional</span><span class="sxs-lookup"><span data-stu-id="6461c-118">optional</span></span>|<span data-ttu-id="6461c-p101">Especifica a versão mínima do conjunto de APIs exigido pelo seu suplemento. Substitui o valor de **DefaultMinVersion**, se ele estiver especificado no elemento [Sets](sets.md) pai.</span><span class="sxs-lookup"><span data-stu-id="6461c-p101">Specifies the minimum version of the API set required by your add-in. Overrides the value of  **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="6461c-121">Comentários</span><span class="sxs-lookup"><span data-stu-id="6461c-121">Remarks</span></span>

<span data-ttu-id="6461c-122">Para obter mais informações sobre os conjuntos de requisitos, consulte [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="6461c-122">For more information about available requirement sets, see [Office add-in requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="6461c-123">Para saber mais sobre o atributo **MinVersion** do elemento **Set** e o atributo **DefaultMinVersion** do elemento **Sets**, confira [Definir o elemento Requirements no manifesto](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="6461c-123">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="6461c-124">Para suplementos de email, há somente um `"Mailbox"` conjunto de requisitos disponível.</span><span class="sxs-lookup"><span data-stu-id="6461c-124">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="6461c-125">Para suplementos de email, há somente um conjunto de requisitos  contenteditable="false" class="locked monad selfClosingTag">`"Mailbox"` disponível.</span><span class="sxs-lookup"><span data-stu-id="6461c-125">Important  For mail add-ins, there is only one   requirement set available. This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins). Also, you can't declare support for specific methods in mail add-ins.</span></span> <span data-ttu-id="6461c-126">Além disso, você não pode declarar suporte para métodos específicos em suplementos de email.</span><span class="sxs-lookup"><span data-stu-id="6461c-126">Also, you can't declare support for specific methods in mail add-ins.</span></span>
