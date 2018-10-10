# <a name="page-element"></a><span data-ttu-id="d064e-101">Elemento Page</span><span class="sxs-lookup"><span data-stu-id="d064e-101">Page element</span></span>

<span data-ttu-id="d064e-102">Define as configurações de página HTML usadas por uma função personalizada no Excel.</span><span class="sxs-lookup"><span data-stu-id="d064e-102">Defines HTML page settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="d064e-103">Atributos</span><span class="sxs-lookup"><span data-stu-id="d064e-103">Attributes</span></span>

<span data-ttu-id="d064e-104">Nenhum</span><span class="sxs-lookup"><span data-stu-id="d064e-104">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="d064e-105">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="d064e-105">Child elements</span></span>

|  <span data-ttu-id="d064e-106">Elemento</span><span class="sxs-lookup"><span data-stu-id="d064e-106">Element</span></span>  |  <span data-ttu-id="d064e-107">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="d064e-107">Required</span></span>  |  <span data-ttu-id="d064e-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="d064e-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="d064e-109">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="d064e-109">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="d064e-110">Sim</span><span class="sxs-lookup"><span data-stu-id="d064e-110">Yes</span></span>  | <span data-ttu-id="d064e-111">ID do recurso do arquivo HTML usado por funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="d064e-111">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="d064e-112">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d064e-112">Example</span></span>

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
