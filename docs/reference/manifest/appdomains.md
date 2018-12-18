# <a name="appdomains-element"></a><span data-ttu-id="d224f-101">Elemento AppDomains</span><span class="sxs-lookup"><span data-stu-id="d224f-101">AppDomains element</span></span>

<span data-ttu-id="d224f-p101">Lista qualquer domínio além do domínio especificado no elemento SourceLocation que seu Suplemento do Office utilizará para carregar páginas. Para cada domínio adicional, especifique um elemento AppDomain.</span><span class="sxs-lookup"><span data-stu-id="d224f-p101">Lists any domains in addition to the domain specified in the SourceLocation element that your Office Add-in will use to load pages. For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="d224f-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="d224f-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="d224f-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="d224f-105">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> <span data-ttu-id="d224f-106">O valor de cada elemento **AppDomain** deve incluir o protocolo (por exemplo, `<AppDomain>https://myappdomain<AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="d224f-106">The value of each **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="d224f-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="d224f-107">Contained in</span></span>

[<span data-ttu-id="d224f-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="d224f-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="d224f-109">Pode conter</span><span class="sxs-lookup"><span data-stu-id="d224f-109">Can contain</span></span>

[<span data-ttu-id="d224f-110">AppDomain</span><span class="sxs-lookup"><span data-stu-id="d224f-110">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="d224f-111">Comentários</span><span class="sxs-lookup"><span data-stu-id="d224f-111">Remarks</span></span>

<span data-ttu-id="d224f-112">Por padrão, o seu suplemento pode carregar qualquer página que esteja no mesmo domínio que o local especificado no elemento [SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="d224f-112">By default, your add-in can load any page that is in the same domain as the location specified in the SourceLocation element. To load pages that are not in the same domain as the add-in, specify the domains by using the AppDomains and AppDomain elements. This element can't be empty.</span></span> <span data-ttu-id="d224f-113">Para carregar páginas que não estejam no mesmo domínio do que o suplemento, especifique os domínios usando os elementos **AppDomains** e **AppDomain**.</span><span class="sxs-lookup"><span data-stu-id="d224f-113">To load pages that are not in the same domain as the add-in, specify their domains by using the **AppDomains** and **AppDomain** elements.</span></span> <span data-ttu-id="d224f-114">Esse elemento não pode estar vazio.</span><span class="sxs-lookup"><span data-stu-id="d224f-114">This element can't be empty.</span></span>
