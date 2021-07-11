<span data-ttu-id="c78ea-101">Outlook os complementos usam principalmente as APIs expostas por meio do [objeto Mailbox.](/javascript/api/outlook/office.mailbox)</span><span class="sxs-lookup"><span data-stu-id="c78ea-101">Outlook add-ins primarily use the APIs exposed through the [Mailbox](/javascript/api/outlook/office.mailbox) object.</span></span> <span data-ttu-id="c78ea-102">Para acessar os objetos e membros específicos para suplementos do Outlook, como o objeto [Item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md), use a propriedade [mailbox](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md) do objeto **Context** para acessar o objeto **Mailbox**, conforme exibido na linha de código abaixo.</span><span class="sxs-lookup"><span data-stu-id="c78ea-102">To access the objects and members specifically for use in Outlook add-ins, such as the [Item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) object, you use the [mailbox](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md) property of the **Context** object to access the **Mailbox** object, as shown in the following line of code.</span></span>

```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

<span data-ttu-id="c78ea-103">Além disso, Outlook os complementos podem usar os seguintes objetos.</span><span class="sxs-lookup"><span data-stu-id="c78ea-103">Additionally, Outlook add-ins can use the following objects.</span></span>

-  <span data-ttu-id="c78ea-104">Objeto **Office**: para inicialização.</span><span class="sxs-lookup"><span data-stu-id="c78ea-104">**Office** object: for initialization.</span></span>

-  <span data-ttu-id="c78ea-105">Objeto **Context**: para acesso a propriedades de conteúdo e idioma de exibição.</span><span class="sxs-lookup"><span data-stu-id="c78ea-105">**Context** object: for access to content and display language properties.</span></span>

-  <span data-ttu-id="c78ea-106">Objeto **RoamingSettings**: para salvar as configurações personalizadas do suplemento do Outlook na caixa de correio do usuário em que o suplemento está instalado.</span><span class="sxs-lookup"><span data-stu-id="c78ea-106">**RoamingSettings** object: for saving Outlook add-in-specific custom settings to the user's mailbox where the add-in is installed.</span></span>

<span data-ttu-id="c78ea-107">Para obter informações sobre como usar a OUTLOOK JavaScript, [consulte Outlook de complementos](../outlook/outlook-add-ins-overview.md).</span><span class="sxs-lookup"><span data-stu-id="c78ea-107">For information about using the Outlook JavaScript API, see [Outlook add-ins](../outlook/outlook-add-ins-overview.md).</span></span>