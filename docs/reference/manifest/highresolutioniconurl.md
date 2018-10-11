# <a name="highresolutioniconurl-element"></a><span data-ttu-id="6b5e8-101">Elemento HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="6b5e8-101">HighResolutionIconUrl element</span></span>

<span data-ttu-id="6b5e8-102">Especifica a URL da imagem usada para representar seu suplemento do Office no UX e no Office Store de inserção em telas de alto DPI.</span><span class="sxs-lookup"><span data-stu-id="6b5e8-102">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store on high DPI screens.</span></span>

<span data-ttu-id="6b5e8-103">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="6b5e8-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="6b5e8-104">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="6b5e8-104">Syntax</span></span>

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="6b5e8-105">Pode conter</span><span class="sxs-lookup"><span data-stu-id="6b5e8-105">Can contain:</span></span>

[<span data-ttu-id="6b5e8-106">Substituição</span><span class="sxs-lookup"><span data-stu-id="6b5e8-106">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="6b5e8-107">Atributos</span><span class="sxs-lookup"><span data-stu-id="6b5e8-107">Attributes</span></span>

|<span data-ttu-id="6b5e8-108">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="6b5e8-108">**Attribute**</span></span>|<span data-ttu-id="6b5e8-109">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="6b5e8-109">**Type**</span></span>|<span data-ttu-id="6b5e8-110">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="6b5e8-110">**Required**</span></span>|<span data-ttu-id="6b5e8-111">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="6b5e8-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="6b5e8-112">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="6b5e8-112">DefaultValue</span></span>|<span data-ttu-id="6b5e8-113">string (URL)</span><span class="sxs-lookup"><span data-stu-id="6b5e8-113">string (URL)</span></span>|<span data-ttu-id="6b5e8-114">obrigatório</span><span class="sxs-lookup"><span data-stu-id="6b5e8-114">required</span></span>|<span data-ttu-id="6b5e8-115">Especifica o valor padrão para essa configuração, expresso para a localidade especificada no elemento [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="6b5e8-115">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="6b5e8-116">Comentários</span><span class="sxs-lookup"><span data-stu-id="6b5e8-116">Remarks</span></span>

<span data-ttu-id="6b5e8-p101">Para um suplemento de email, o ícone é exibido na interface de usuário **Arquivo**  >  **Gerenciar suplementos**. Para um suplemento de conteúdo ou de painel de tarefas, o ícone é exibido na interface de usuário **Inserir**  >  **Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="6b5e8-p101">For a mail add-in, the icon is displayed in the  **File** > **Manage add-ins** UI . For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span>

<span data-ttu-id="6b5e8-119">A imagem deve estar em um dos seguintes formatos de arquivo em uma resolução recomendada de 64 x 64 pixels: GIF, JPG, PNG, EXIF, BMP ou TIFF.</span><span class="sxs-lookup"><span data-stu-id="6b5e8-119">The image must be in one of the following file formats at a recommended resolution of 64 x 64 pixels: GIF, JPG, PNG, EXIF, BMP or TIFF.</span></span> <span data-ttu-id="6b5e8-120">Para obter mais informações, consulte a seção _Criar uma identidade visual consistente para seu aplicativo_ em [Criar listagens eficazes na AppSource e no Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings).</span><span class="sxs-lookup"><span data-stu-id="6b5e8-120">For more information, see the section  Create a consistent visual identity for your app in Create effective Office Store apps and add-ins.</span></span>
