# <a name="appdomains-element"></a><span data-ttu-id="a197a-101">Elemento AppDomains</span><span class="sxs-lookup"><span data-stu-id="a197a-101">AppDomains element</span></span>

<span data-ttu-id="a197a-p101">Lista qualquer domínio além do domínio especificado no elemento SourceLocation que seu Suplemento do Office utilizará para carregar páginas. Para cada domínio adicional, especifique um elemento AppDomain.</span><span class="sxs-lookup"><span data-stu-id="a197a-p101">Lists any domains in addition to the domain specified in the SourceLocation element that your Office Add-in will use to load pages. For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="a197a-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas, E-mail</span><span class="sxs-lookup"><span data-stu-id="a197a-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="a197a-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="a197a-105">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

## <a name="contained-in"></a><span data-ttu-id="a197a-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="a197a-106">Contained in:</span></span>

[<span data-ttu-id="a197a-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="a197a-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="a197a-108">Pode conter</span><span class="sxs-lookup"><span data-stu-id="a197a-108">Can contain:</span></span>

[<span data-ttu-id="a197a-109">AppDomain</span><span class="sxs-lookup"><span data-stu-id="a197a-109">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="a197a-110">Comentários</span><span class="sxs-lookup"><span data-stu-id="a197a-110">Remarks</span></span>

<span data-ttu-id="a197a-p102">Por padrão, o seu suplemento pode carregar qualquer página que esteja no mesmo domínio que o local especificado no elemento **SourceLocation**. Para carregar páginas que não estejam no mesmo domínio que o suplemento, especifique os domínios usando os elementos **AppDomains** e **AppDomain**. Esse elemento não pode estar vazio.</span><span class="sxs-lookup"><span data-stu-id="a197a-p102">By default, your add-in can load any page that is in the same domain as the location specified in the **SourceLocation** element. To load pages that are not in the same domain as the add-in, specify the domains by using the **AppDomains** and **AppDomain** elements. This element can't be empty.</span></span> 
