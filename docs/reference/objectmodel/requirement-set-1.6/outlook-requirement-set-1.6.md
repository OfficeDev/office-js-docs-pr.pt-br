# <a name="outlook-add-in-api-requirement-set-16"></a>Conjunto de requisitos de API para suplementos do Outlook versão 1.6

O subconjunto de APIs para suplementos do Outlook da API JavaScript para Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.

> [!NOTE]
> Esta documentação destina-se a um [conjunto de requisitos](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) diferente do conjunto de requisitos mais recente.

## <a name="whats-new-in-16"></a>Novidades na versão 1.6

O conjunto de requisitos 1.6 inclui todos os recursos do [Conjunto de requisitos 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md), além dos seguintes.

- Adicionadas novas APIs para suplementos contextuais para obter a equivalência de entidade ou RegEx que o usuário selecionou para ativar o suplemento.
- Adicionada uma nova API para abrir um novo formulário de mensagem.
- Adicionada a capacidade para o suplemento determinar o tipo de conta da caixa de correio do usuário.

### <a name="change-log"></a>Log de alterações

- Adicionado [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#getselectedentities--entitiesjavascriptapioutlook16officeentities): Adiciona uma nova função que obtém as entidades encontradas em uma correspondência destacada selecionada pelo usuário. Correspondências destacadas se aplicam aos suplementos contextuais.
- Adicionado [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#getselectedregexmatches--object): Adiciona uma nova função que retorna valores de sequência de caracteres em uma correspondência destacada que corresponde às expressões regulares definidas no arquivo de manifesto XML. Correspondências destacadas se aplicam aos suplementos contextuais.
- Adicionado [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#displaynewmessageformparameters): Adiciona uma nova função que abre um novo formulário de mensagem.
- Adicionada [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#accounttype-string): adiciona um novo membro para o perfil de usuário que indica o tipo de conta do usuário.

## <a name="see-also"></a>Consulte também

- [Suplementos do Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](https://docs.microsoft.com/outlook/add-ins/quick-start)