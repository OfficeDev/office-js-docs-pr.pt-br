 

# <a name="office"></a>Office

O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

##### <a name="members-and-methods"></a>Membros e métodos

| Membro | Tipo |
|--------|------|
| [AsyncResultStatus](#asyncresultstatus-string) | Membro |
| [CoercionType](#coerciontype-string) | Membro |
| [EventType](#eventtype-string) | Membro |
| [SourceProperty](#sourceproperty-string) | Membro |

### <a name="namespaces"></a>Namespaces

[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API dos suplementos do Office para uso na API do suplemento do Outlook.

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.

### <a name="members"></a>Membros

####  <a name="asyncresultstatus-string"></a>AsyncResultStatus :String

Especifica o resultado de uma chamada assíncrona.

##### <a name="type"></a>Tipo:

*   Cadeia de caracteres

##### <a name="properties"></a>Propriedades:

|Nome| Tipo| Descrição|
|---|---|---|
|`Succeeded`| Cadeia de caracteres|A chamada foi bem-sucedida.|
|`Failed`| Cadeia de caracteres|A chamada falhou.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

---

####  <a name="coerciontype-string"></a>CoercionType :String

Especifica como forçar os dados retornados ou definir de acordo com o método invocado.

##### <a name="type"></a>Tipo:

*   Cadeia de caracteres

##### <a name="properties"></a>Propriedades:

|Nome| Tipo| Descrição|
|---|---|---|
|`Html`| Cadeia de caracteres|Solicita que os dados sejam retornados no formato HTML.|
|`Text`| Cadeia de caracteres|Solicita que os dados sejam retornados no formato de texto.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo aplicável ao Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

---

####  <a name="eventtype-string"></a>EventType :String

Especifica o evento associado a um manipulador de eventos.

##### <a name="type"></a>Tipo:

*   Cadeia de caracteres

##### <a name="properties"></a>Propriedades:

| Nome | Tipo | Descrição | Conjunto de requisitos mínimos |
|---|---|---|---|
|`AppointmentTimeChanged`| Cadeia de caracteres | A data ou hora do compromisso selecionado ou série foi alterada. | 1.7 |
|`ItemChanged`| Cadeia de caracteres | O item selecionado foi alterado. | 1.5 |
|`OfficeThemeChanged`| Cadeia de caracteres | O item selecionado foi alterado. | Visualizar |
|`RecipientsChanged`| Cadeia de caracteres | A lista de destinatários do item selecionado ou o local do compromisso foram alterados. | 1.7 |
|`RecurrenceChanged`| Cadeia de caracteres | O padrão de recorrência da série selecionada foi alterado. | 1.7 |

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler |

---

####  <a name="sourceproperty-string"></a>SourceProperty :String

Especifica a origem dos dados retornados pelo método invocado.

##### <a name="type"></a>Tipo:

*   Cadeia de caracteres

##### <a name="properties"></a>Propriedades:

|Nome| Tipo| Descrição|
|---|---|---|
|`Body`| Cadeia de caracteres|A origem dos dados é do corpo de uma mensagem.|
|`Subject`| Cadeia de caracteres|A origem dos dados é do assunto de uma mensagem.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo aplicável ao Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|