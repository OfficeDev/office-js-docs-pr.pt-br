Os suplementos do Outlook usam principalmente um subconjunto da API exposta no objeto [Mailbox](/javascript/api/outlook/office.mailbox). Para acessar os objetos e membros especificamente para uso em suplementos do Outlook, como o [objeto Item](/javascript/api/outlook/office.item), use a propriedade de [](/javascript/api/office/office.context#office-office-context-mailbox-member) caixa de correio do objeto **Context** para acessar o objeto **Mailbox**, conforme mostrado na linha de código a seguir.

```js
// Access the Item object.
const item = Office.context.mailbox.item;
```

Além disso, os suplementos do Outlook podem usar os objetos a seguir.

- Objeto **Office**: para inicialização.

- Objeto **Context**: para acesso a propriedades de conteúdo e idioma de exibição.

- Objeto **RoamingSettings**: para salvar as configurações personalizadas do suplemento do Outlook na caixa de correio do usuário em que o suplemento está instalado.

Para obter informações sobre como usar o JavaScript em suplementos do Outlook, confira [Suplementos do Outlook ](../outlook/outlook-add-ins-overview.md).
