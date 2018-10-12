# <a name="outlook-add-in-api-requirement-set-14"></a><span data-ttu-id="b1372-101">Conjunto de requisitos de API para suplementos do Outlook versão 1.4</span><span class="sxs-lookup"><span data-stu-id="b1372-101">Outlook add-in API requirement set 1.4</span></span>

<span data-ttu-id="b1372-102">O subconjunto da API para suplementos do Outlook da API JavaScript para Office para inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="b1372-102">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="b1372-103">Esta documentação destina-se a um [conjunto de requisitos](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) diferente do conjunto de requisitos mais recente.</span><span class="sxs-lookup"><span data-stu-id="b1372-103">Note: This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-14"></a><span data-ttu-id="b1372-104">Novidades na versão 1.4?</span><span class="sxs-lookup"><span data-stu-id="b1372-104">What's new in 1.4?</span></span>

<span data-ttu-id="b1372-p101">O conjunto de requisitos versão 1.4 inclui todos os recursos do [Conjunto de requisitos versão 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). O acesso ao namespace `Office.ui` foi adicionado.</span><span class="sxs-lookup"><span data-stu-id="b1372-p101">Requirement set 1.4 includes all of the features of [Requirement set 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). It added access to the `Office.ui` namespace.</span></span>

### <a name="change-log"></a><span data-ttu-id="b1372-107">Log de alterações</span><span class="sxs-lookup"><span data-stu-id="b1372-107">Change log</span></span>

- <span data-ttu-id="b1372-108">[Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) adicionado: exibe uma caixa de diálogo em um host do Office.</span><span class="sxs-lookup"><span data-stu-id="b1372-108">Added [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-): Displays a dialog box in an Office host.</span></span>
- <span data-ttu-id="b1372-109">[Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-messageobject-) adicionado: fornece uma mensagem da caixa de diálogo à sua página pai/de abertura.</span><span class="sxs-lookup"><span data-stu-id="b1372-109">Added [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-messageobject-): Delivers a message from the dialog box to its parent/opener page.</span></span>
- <span data-ttu-id="b1372-110">Objeto [Dialog](/javascript/api/office/office.dialog) adicionado: o objeto retornado quando o método  [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) é chamado.</span><span class="sxs-lookup"><span data-stu-id="b1372-110">Added Dialog object: The object that is returned when the  method is called.</span></span>

## <a name="see-also"></a><span data-ttu-id="b1372-111">Confira também</span><span class="sxs-lookup"><span data-stu-id="b1372-111">See also</span></span>

- [<span data-ttu-id="b1372-112">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="b1372-112">Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/)
- [<span data-ttu-id="b1372-113">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="b1372-113">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="b1372-114">Introdução</span><span class="sxs-lookup"><span data-stu-id="b1372-114">Get started</span></span>](https://docs.microsoft.com/outlook/add-ins/quick-start)