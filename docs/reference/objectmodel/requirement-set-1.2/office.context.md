---
title: Office. Context – conjunto de requisitos 1,2
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: f5b1e610814e5387c6b27c17569b29860bd2e4be
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814966"
---
# <a name="context"></a>context

### <a name="officeofficemdcontext"></a>[Office](office.md).context

O Office. Context fornece interfaces compartilhadas usadas por suplementos em todos os aplicativos do Office. Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-1.2).

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

##### <a name="properties"></a>Propriedades

| Propriedade | Modelos | Tipo de retorno | Mínimo<br>conjunto de requisitos |
|---|---|---|:---:|
| [contentLanguage](#contentlanguage-string) | Escrever<br>Leitura | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [la](#diagnostics-contextinformation) | Escrever<br>Leitura | [ContextInformation](/javascript/api/office/office.contextinformation?view=outlook-js-1.2) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayLanguage](#displaylanguage-string) | Escrever<br>Leitura | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [principal](#host-hosttype) | Escrever<br>Leitura | [HostType](/javascript/api/office/office.hosttype?view=outlook-js-1.2) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [mailbox](office.context.mailbox.md) | Escrever<br>Leitura | [Caixa de Correio](/javascript/api/office/office.mailbox?view=outlook-js-1.2) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [plataforma](#platform-platformtype) | Escrever<br>Leitura | [PlatformType](/javascript/api/office/office.platformtype?view=outlook-js-1.2) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [atende](#requirements-requirementsetsupport) | Escrever<br>Leitura | [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.2) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [roamingSettings](#roamingsettings-roamingsettings) | Escrever<br>Leitura | [RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.2) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ui](#ui-ui) | Escrever<br>Leitura | [UI](/javascript/api/office/office.ui?view=outlook-js-1.2) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a>Detalhes da propriedade

#### <a name="contentlanguage-string"></a>contentLanguage: cadeia de caracteres

Obtém a localidade (idioma) especificada pelo usuário para edição do item.

O `contentLanguage` valor reflete a configuração de **idioma de edição** atual especificada com opções de **arquivo > > idioma** no aplicativo host do Office.

##### <a name="type"></a>Tipo

*   String

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

##### <a name="example"></a>Exemplo

```js
function sayHelloWithContentLanguage() {
  var myContentLanguage = Office.context.contentLanguage;
  switch (myContentLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}

// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

<br>

---
---

#### <a name="diagnostics-contextinformationjavascriptapiofficeofficecontextinformation"></a>diagnóstico: [ContextInformation](/javascript/api/office/office.contextinformation)

Obtém informações sobre o ambiente no qual o suplemento está sendo executado.

##### <a name="type"></a>Tipo

*   [ContextInformation](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

##### <a name="example"></a>Exemplo

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a>displayLanguage: cadeia de caracteres

Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo host do Office.

O valor `displayLanguage` reflete a configuração atual de **Display Language** especificada com **Arquivo > Opções > Idioma** no aplicativo host do Office.

##### <a name="type"></a>Tipo

*   String

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

##### <a name="example"></a>Exemplo

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}

// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

<br>

---
---

#### <a name="host-hosttypejavascriptapiofficeofficehosttype"></a>host: [HostType](/javascript/api/office/office.hosttype)

Obtém o host do aplicativo do Office no qual o suplemento está sendo executado.

##### <a name="type"></a>Tipo

*   [HostType](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

##### <a name="example"></a>Exemplo

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtypejavascriptapiofficeofficeplatformtype"></a>Platform: [platformtype](/javascript/api/office/office.platformtype)

Fornece a plataforma na qual o suplemento está sendo executado.

##### <a name="type"></a>Tipo

*   [PlatformType](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

##### <a name="example"></a>Exemplo

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupportjavascriptapiofficeofficerequirementsetsupport"></a>requisitos: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)

Fornece um método para determinar quais conjuntos de requisitos são compatíveis com o host atual e a plataforma.

##### <a name="type"></a>Tipo

*   [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

##### <a name="example"></a>Exemplo

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a>roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)

Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.

O objeto `RoamingSettings` permite armazenar e acessar os dados de um suplemento de email que está armazenado na caixa de correio do usuário, para que fiquem disponíveis para esse suplemento quando ele for executado em qualquer aplicativo host de cliente usado para acessar essa caixa de correio.

##### <a name="type"></a>Tipo

*   [RoamingSettings](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| Restrito|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

<br>

---
---

#### <a name="ui-uijavascriptapiofficeofficeui"></a>UI: [UI](/javascript/api/office/office.ui)

Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus suplementos do Office.

##### <a name="type"></a>Tipo

*   [UI](/javascript/api/office/office.ui)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|
