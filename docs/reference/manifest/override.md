# <a name="override-element"></a><span data-ttu-id="4f60a-101">Elemento Override</span><span class="sxs-lookup"><span data-stu-id="4f60a-101">Override element</span></span>

<span data-ttu-id="4f60a-102">Fornece uma maneira de especificar o valor de uma configuração para uma localidade adicional.</span><span class="sxs-lookup"><span data-stu-id="4f60a-102">Provides a way to specify the value of a setting for an additional locale.</span></span>

<span data-ttu-id="4f60a-103">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="4f60a-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="4f60a-104">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="4f60a-104">Syntax</span></span>

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a><span data-ttu-id="4f60a-105">Contido em</span><span class="sxs-lookup"><span data-stu-id="4f60a-105">Contained in:</span></span>

|<span data-ttu-id="4f60a-106">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="4f60a-106">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="4f60a-107">CitationText</span><span class="sxs-lookup"><span data-stu-id="4f60a-107">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="4f60a-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="4f60a-108">Description</span></span>](description.md)|
|[<span data-ttu-id="4f60a-109">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="4f60a-109">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="4f60a-110">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="4f60a-110">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="4f60a-111">DisplayName</span><span class="sxs-lookup"><span data-stu-id="4f60a-111">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="4f60a-112">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="4f60a-112">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="4f60a-113">IconUrl</span><span class="sxs-lookup"><span data-stu-id="4f60a-113">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="4f60a-114">QueryUri</span><span class="sxs-lookup"><span data-stu-id="4f60a-114">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="4f60a-115">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="4f60a-115">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="4f60a-116">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="4f60a-116">SupportUrl</span></span>](supporturl.md)|

## <a name="attributes"></a><span data-ttu-id="4f60a-117">Atributos</span><span class="sxs-lookup"><span data-stu-id="4f60a-117">Attributes</span></span>

|<span data-ttu-id="4f60a-118">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="4f60a-118">**Attribute**</span></span>|<span data-ttu-id="4f60a-119">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="4f60a-119">**Type**</span></span>|<span data-ttu-id="4f60a-120">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="4f60a-120">**Required**</span></span>|<span data-ttu-id="4f60a-121">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="4f60a-121">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="4f60a-122">Localidade</span><span class="sxs-lookup"><span data-stu-id="4f60a-122">Locale</span></span>|<span data-ttu-id="4f60a-123">string</span><span class="sxs-lookup"><span data-stu-id="4f60a-123">string</span></span>|<span data-ttu-id="4f60a-124">required</span><span class="sxs-lookup"><span data-stu-id="4f60a-124">required</span></span>|<span data-ttu-id="4f60a-125">Especifica o nome da cultura da localidade para essa substituição no formato de marca do idioma BCP 47, como `"en-US"`.</span><span class="sxs-lookup"><span data-stu-id="4f60a-125">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="4f60a-126">Valor</span><span class="sxs-lookup"><span data-stu-id="4f60a-126">Value</span></span>|<span data-ttu-id="4f60a-127">string</span><span class="sxs-lookup"><span data-stu-id="4f60a-127">string</span></span>|<span data-ttu-id="4f60a-128">required</span><span class="sxs-lookup"><span data-stu-id="4f60a-128">required</span></span>|<span data-ttu-id="4f60a-129">Especifica o valor da configuração expressa para a localidade especificada.</span><span class="sxs-lookup"><span data-stu-id="4f60a-129">Specifies value of the setting expressed for the specified locale.</span></span>|

## <a name="see-also"></a><span data-ttu-id="4f60a-130">Confira também</span><span class="sxs-lookup"><span data-stu-id="4f60a-130">See also</span></span>

- [<span data-ttu-id="4f60a-131">Localização dos suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="4f60a-131">Localization for Office Add-ins</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/localization)
    
