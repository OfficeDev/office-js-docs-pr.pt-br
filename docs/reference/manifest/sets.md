# <a name="sets-element"></a><span data-ttu-id="52883-101">Elemento Sets</span><span class="sxs-lookup"><span data-stu-id="52883-101">Sets element</span></span>

<span data-ttu-id="52883-102">Especifica o subconjunto mínimo da API JavaScript para Office que o suplemento do Office exige para ativar.</span><span class="sxs-lookup"><span data-stu-id="52883-102">Specifies the minimum subset of the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="52883-103">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="52883-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="52883-104">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="52883-104">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="52883-105">Contido em</span><span class="sxs-lookup"><span data-stu-id="52883-105">Contained in:</span></span>

[<span data-ttu-id="52883-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="52883-106">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="52883-107">Pode conter</span><span class="sxs-lookup"><span data-stu-id="52883-107">Can contain:</span></span>

[<span data-ttu-id="52883-108">Set</span><span class="sxs-lookup"><span data-stu-id="52883-108">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="52883-109">Atributos</span><span class="sxs-lookup"><span data-stu-id="52883-109">Attributes</span></span>

|<span data-ttu-id="52883-110">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="52883-110">**Attribute**</span></span>|<span data-ttu-id="52883-111">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="52883-111">**Type**</span></span>|<span data-ttu-id="52883-112">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="52883-112">**Required**</span></span>|<span data-ttu-id="52883-113">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="52883-113">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="52883-114">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="52883-114">DefaultMinVersion</span></span>|<span data-ttu-id="52883-115">string</span><span class="sxs-lookup"><span data-stu-id="52883-115">string</span></span>|<span data-ttu-id="52883-116">optional</span><span class="sxs-lookup"><span data-stu-id="52883-116">optional</span></span>|<span data-ttu-id="52883-p101">Especifica o valor padrão do atributo  **MinVersion** para todos os elementos [Set](set.md) filhos. O valor padrão é "1.1".</span><span class="sxs-lookup"><span data-stu-id="52883-p101">Specifies the default  **MinVersion** attribute value for all child [Set](set.md) elements. The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="52883-119">Comentários</span><span class="sxs-lookup"><span data-stu-id="52883-119">Remarks</span></span>

<span data-ttu-id="52883-120">Para obter mais informações sobre os conjuntos de requisitos, consulte [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="52883-120">For more information about available requirement sets, see [Office add-in requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="52883-121">Para saber mais sobre o atributo **MinVersion** do elemento **Set** e o atributo **DefaultMinVersion** do elemento **Sets**, confira [Definir o elemento Requirements no manifesto](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="52883-121">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

