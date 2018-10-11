 

# <a name="office"></a>Office

O namespace Office fornece interfaces compartilhadas que são usadas por suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma listagem completa do namespace Office, confira [API compartilhada](/javascript/api/office).

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto mínimo de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

### <a name="namespaces"></a>Namespaces

[context](Office.context.md): fornece interfaces compartilhadas do namespace do contexto da API de suplementos do Office para uso na API de suplemento do Outlook.

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.

### <a name="members"></a>Membros

####  <a name="asyncresultstatus-string"></a>AsyncResultStatus :String

Especifica o resultado de uma chamada assíncrona.

##### <a name="type"></a>Tipo:

*   Sequência de caracteres

##### <a name="properties"></a>Propriedades:

|Nome| Tipo| Descrição|
|---|---|---|
|`Succeeded`| Sequência de caracteres|A chamada foi bem-sucedida.|
|`Failed`| Sequência de caracteres|A chamada falhou.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto mínimo de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|
####  <a name="coerciontype-string"></a>CoercionType :String

Especifica como forçar os dados retornados ou atribuídos pelo método chamado.

##### <a name="type"></a>Tipo:

*   Sequência de caracteres

##### <a name="properties"></a>Propriedades:

|Nome| Tipo| Descrição|
|---|---|---|
|`Html`| Sequência de caracteres|Solicita que os dados sejam retornados no formato HTML.|
|`Text`| Sequência de caracteres|Solicita que os dados sejam retornados no formato de texto.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto mínimo de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|
####  <a name="sourceproperty-string"></a>SourceProperty :String

Especifica a origem dos dados retornados pelo método chamado.

##### <a name="type"></a>Tipo:

*   Sequência de caracteres

##### <a name="properties"></a>Propriedades:

|Nome| Tipo| Descrição|
|---|---|---|
|`Body`| Sequência de caracteres|A origem dos dados é do corpo de uma mensagem.|
|`Subject`| Sequência de caracteres|A origem dos dados é do assunto de uma mensagem.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto mínimo de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|