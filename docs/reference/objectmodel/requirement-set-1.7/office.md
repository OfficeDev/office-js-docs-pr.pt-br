 

# <a name="office"></a>Office

O namespace Office fornece interfaces compartilhadas que são usadas por suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma listagem completa do namespace Office, confira [API compartilhada](/javascript/api/office).

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

##### <a name="members-and-methods"></a>Membros e métodos

| Membro | Tipo |
|--------|------|
| [AsyncResultStatus](#asyncresultstatus-string) | Membro |
| [CoercionType](#coerciontype-string) | Membro |
| [EventType](#eventtype-string) | Membro |
| [SourceProperty](#sourceproperty-string) | Membro |

### <a name="namespaces"></a>Namespaces

[context](office.context.md): fornece interfaces compartilhadas do namespace do contexto da API de suplementos do Office para uso na API de suplemento do Outlook.

[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.

### <a name="members"></a>Membros

####  <a name="asyncresultstatus-string"></a>AsyncResultStatus :String

Especifica o resultado de uma chamada assíncrona.

##### <a name="type"></a>Tipo:

*   sequência de caracteres

##### <a name="properties"></a>Propriedades:

|Nome| Tipo| Descrição|
|---|---|---|
|`Succeeded`| sequência de caracteres|A chamada foi bem-sucedida.|
|`Failed`| sequência de caracteres|A chamada falhou.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

---

####  <a name="coerciontype-string"></a>CoercionType :String

Especifica como forçar os dados retornados ou definidos de acordo com o método invocado.

##### <a name="type"></a>Tipo:

*   sequência de caracteres

##### <a name="properties"></a>Propriedades:

|Nome| Tipo| Descrição|
|---|---|---|
|`Html`| sequência de caracteres|Solicita que os dados sejam retornados no formato HTML.|
|`Text`| sequência de caracteres|Solicita que os dados sejam retornados no formato de texto.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

---

####  <a name="eventtype-string"></a>EventType :String

Especifica o evento associado a um manipulador de eventos.

##### <a name="type"></a>Tipo:

*   sequência de caracteres

##### <a name="properties"></a>Propriedades:

| Nome | Tipo | Descrição | Conjunto de requisitos mínimos |
|---|---|---|---|
|`AppointmentTimeChanged`| sequência de caracteres | A data ou hora do compromisso selecionado ou série foi alterada. | 1.7 |
|`ItemChanged`| sequência de caracteres | O item selecionado foi alterado. | 1.5 |
|`RecipientsChanged`| sequência de caracteres | A lista de destinatários do item ou local do compromisso selecionado foram alterados. | 1.7 |
|`RecurrenceChanged`| sequência de caracteres | O padrão de recorrência da série selecionada foi alterado. | 1.7 |

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler |

---

####  <a name="sourceproperty-string"></a>SourceProperty :String

Especifica a origem dos dados retornados pelo método invocado.

##### <a name="type"></a>Tipo:

*   sequência de caracteres

##### <a name="properties"></a>Propriedades:

|Nome| Tipo| Descrição|
|---|---|---|
|`Body`| sequência de caracteres|A origem dos dados é do corpo de uma mensagem.|
|`Subject`| sequência de caracteres|A origem dos dados é do assunto de uma mensagem.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|