# <a name="outlook-add-in-api-preview-requirement-set"></a>Conjunto de requisitos da API de suplementos do Outlook em versão prévia

O subconjunto da API JavaScript para Office  da API para suplementos do Outlook inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.

> [!NOTE]
> Esta documentação refere-se à **versão prévia** do [conjunto de requisitos](/javascript/office/requirement-sets/outlook-api-requirement-sets). Esse conjunto de requisitos ainda não está totalmente implementado e os clientes não informarão precisamente o suporte para ele. Você não deve especificar a esse conjunto de requisitos em seu manifesto de suplemento. Os métodos e propriedades introduzidos neste conjunto de requisitos devem ser testados individualmente para verificar sua disponibilidade antes de usá-los.

O conjunto de requisitos em versão prévia inclui todos os recursos do [Conjunto de requisitos 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).

## <a name="features-in-preview"></a>Recursos em versão prévia

Os seguintes recursos estão em versão prévia.

- [SharedProperties](/javascript/api/outlook/office.sharedproperties) - foi adicionado um novo objeto que representa as propriedades de um item de compromisso ou mensagem em uma pasta, calendário ou caixa de correio compartilhada.
- [Event.completed](/javascript/api/office/office.addincommands.event#completed-options-) - um novo parâmetro opcional `options`  que é um dicionário com um valor válido `allowEvent`. Esse valor é usado para cancelar a execução de um evento.
- [Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) - foi adicionado um novo método que anexa um arquivo da codificação base64 em uma mensagem ou um compromisso.
- [Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback) – foi adicionada uma nova função que retorna os dados de inicialização que são passados quando o suplemento é [ativado por uma mensagem acionável](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).
- [Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback) - foi adicionado um novo método que obtém um objeto que representa as sharedProperties de um item de compromisso ou mensagem.
- [Office.context.auth.getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) – Foi adicionado acesso ao `getAccessTokenAsync`, que permite que os suplementos [obtenham um token de acesso](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) da API do Microsoft Graph.
- [Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) - foi adicionada uma nova enumeração de sinalizador que especifica as permissões de representante.
- [Office.EventType](/javascript/api/office/office.eventtype) - foi modificado para dar suporte ao evento OfficeThemeChanged por meio da adição da entrada `OfficeThemeChanged`.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](https://docs.microsoft.com/outlook/add-ins/quick-start)