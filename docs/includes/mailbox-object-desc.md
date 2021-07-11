Outlook os complementos usam principalmente as APIs expostas por meio do [objeto Mailbox.](/javascript/api/outlook/office.mailbox) Para acessar os objetos e membros específicos para suplementos do Outlook, como o objeto [Item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md), use a propriedade [mailbox](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md) do objeto **Context** para acessar o objeto **Mailbox**, conforme exibido na linha de código abaixo.

```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

Além disso, Outlook os complementos podem usar os seguintes objetos.

-  Objeto **Office**: para inicialização.

-  Objeto **Context**: para acesso a propriedades de conteúdo e idioma de exibição.

-  Objeto **RoamingSettings**: para salvar as configurações personalizadas do suplemento do Outlook na caixa de correio do usuário em que o suplemento está instalado.

Para obter informações sobre como usar a OUTLOOK JavaScript, [consulte Outlook de complementos](../outlook/outlook-add-ins-overview.md).