# <a name="outlook-add-in-api-requirement-set-11"></a>Conjunto de requisitos de API versão 1.1 para suplementos do Outlook

O subconjunto de APIs para suplementos do Outlook das APIs JavaScript para Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.

> [!NOTE]
> Esta documentação destina-se a um [conjunto de requisitos](/javascript/office/requirement-sets/outlook-api-requirement-sets) que não seja o conjunto de requisitos mais recente. 

## <a name="whats-new-in-11"></a>Novidades na versão 1.1?

O conjunto de requisitos 1.1 inclui todos os recursos do conjunto de requisitos 1.0. Ele adicionou a capacidade de os suplementos acessarem o corpo de mensagens e compromissos e a capacidade de modificar o item atual.

### <a name="change-log"></a>Log de alterações

- Foi adicionado o objeto [Body](/javascript/api/outlook_1_1/office.body): fornece métodos para adicionar e atualizar o conteúdo de um item em um suplemento do Outlook.
- Foi adicionado o objeto [Location](/javascript/api/outlook_1_1/office.location): fornece métodos para obter e definir a localização de uma reunião em um suplemento do Outlook.
- Foi adicionado o objeto [Recipients](/javascript/api/outlook_1_1/office.recipients): fornece métodos para obter e definir os destinatários de um compromisso ou mensagem em um suplemento do Outlook.
- Foi adicionado o objeto [Subject](/javascript/api/outlook_1_1/office.subject): fornece métodos para obter e definir o assunto de um compromisso ou mensagem em um suplemento do Outlook.
- Foi adicionado o objeto [Time](/javascript/api/outlook_1_1/office.time): fornece métodos para obter e definir o horário de início ou fim de uma reunião em um suplemento do Outlook.
- Foi adicionado o [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback): adiciona um arquivo a uma mensagem ou compromisso como um anexo.
- Foi adicionado o [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#additemattachmentasyncitemid-attachmentname-options-callback): adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.
- Foi adicionado o [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback): remove um anexo de uma mensagem ou de um compromisso.
- Foi adicionado o [Office.context.mailbox.item.body](office.context.mailbox.item.md#body-bodyjavascriptapioutlook11officebody): obtém um objeto que fornece métodos para manipular o corpo de um item.
- Foi adicionado o [Office.context.mailbox.item.bcc](office.context.mailbox.item.md#bcc-recipientsjavascriptapioutlook11officerecipients): obtém ou define os destinatários na linha Cco (com cópia oculta) de uma mensagem.
- Foi adicionado o [Office.MailboxEnums.RecipientType](/javascript/api/outlook_1_1/office.mailboxenums.recipienttype): especifica o tipo de destinatário para um compromisso.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](https://docs.microsoft.com/outlook/add-ins/quick-start)