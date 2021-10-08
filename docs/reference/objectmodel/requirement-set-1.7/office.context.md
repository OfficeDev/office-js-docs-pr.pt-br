---
title: Office.context - conjunto de requisitos 1.7
description: Office. Membros do objeto context disponível para Outlook os complementos usando o conjunto de requisitos da API de Caixa de Correio 1.7.
ms.date: 12/03/2020
ms.localizationpriority: medium
ms.openlocfilehash: edfc648a7b7e2a7857b2637e2a817c357d9f62c2
ms.sourcegitcommit: efd0966f6400c8e685017ce0c8c016a2cbab0d5c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/08/2021
ms.locfileid: "60237325"
---
# <a name="context-mailbox-requirement-set-17"></a>context (Conjunto de requisitos de caixa de correio 1.7)

### <a name="officecontext"></a>[Office](office.md).context

Office.context fornece interfaces compartilhadas que são usadas por complementos em todos os Office aplicativos. Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-1.7&preserve-view=true).

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Modo do Outlook aplicável](../../../outlook/outlook-add-ins-overview.md#extension-points)| Escrever ou Ler|

## <a name="properties"></a>Propriedades

| Propriedade | Modos | Tipo de retorno | Minimum<br>conjunto de requisitos |
|---|---|---|:---:|
| [contentLanguage](#contentlanguage-string) | Escrever<br>Leitura | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [diagnostics](#diagnostics-contextinformation) | Escrever<br>Leitura | [ContextInformation](/javascript/api/office/office.contextinformation?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayLanguage](#displaylanguage-string) | Escrever<br>Leitura | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [host](#host-hosttype) | Escrever<br>Leitura | [HostType](/javascript/api/office/office.hosttype?view=outlook-js-1.7&preserve-view=true) | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [mailbox](office.context.mailbox.md) | Escrever<br>Leitura | [Caixa de Correio](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [platform](#platform-platformtype) | Escrever<br>Leitura | [PlatformType](/javascript/api/office/office.platformtype?view=outlook-js-1.7&preserve-view=true) | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [requirements](#requirements-requirementsetsupport) | Escrever<br>Leitura | [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [roamingSettings](#roamingsettings-roamingsettings) | Escrever<br>Leitura | [RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ui](#ui-ui) | Escrever<br>Leitura | [UI](/javascript/api/office/office.ui?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a>Detalhes da propriedade

#### <a name="contentlanguage-string"></a>contentLanguage: String

Obtém a localidade (idioma) especificada pelo usuário para editar o item.

O `contentLanguage` valor reflete a **configuração** atual de Idioma de Edição especificada com opções de > de arquivo **> idioma** no aplicativo Office cliente.

##### <a name="type"></a>Tipo

*   String

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Modo do Outlook aplicável](../../../outlook/outlook-add-ins-overview.md#extension-points)| Escrever ou Ler|

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

#### <a name="diagnostics-contextinformation"></a>diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation?view=outlook-js-1.7&preserve-view=true)

Obtém informações sobre o ambiente no qual o complemento está sendo executado.

##### <a name="type"></a>Tipo

*   [ContextInformation](/javascript/api/office/office.contextinformation?view=outlook-js-1.7&preserve-view=true)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Modo do Outlook aplicável](../../../outlook/outlook-add-ins-overview.md#extension-points)| Escrever ou Ler|

##### <a name="example"></a>Exemplo

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a>displayLanguage: String

Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo cliente Office cliente.

O valor reflete a configuração atual de Idioma de Exibição especificada com Opções > > Idioma no aplicativo Office `displayLanguage` cliente.  

##### <a name="type"></a>Tipo

*   String

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Modo do Outlook aplicável](../../../outlook/outlook-add-ins-overview.md#extension-points)| Escrever ou Ler|

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

#### <a name="host-hosttype"></a>host: [HostType](/javascript/api/office/office.hosttype?view=outlook-js-1.7&preserve-view=true)

Obtém o Office aplicativo que está hospedando o complemento.

> [!NOTE]
> Como alternativa, você pode usar a [propriedade Office.context.diagnostics](#diagnostics-contextinformation) para obter o host.

##### <a name="type"></a>Tipo

*   [HostType](/javascript/api/office/office.hosttype?view=outlook-js-1.7&preserve-view=true)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1,5|
|[Modo do Outlook aplicável](../../../outlook/outlook-add-ins-overview.md#extension-points)| Escrever ou Ler|

##### <a name="example"></a>Exemplo

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a>plataforma: [PlatformType](/javascript/api/office/office.platformtype?view=outlook-js-1.7&preserve-view=true)

Fornece a plataforma na qual o complemento está sendo executado.

> [!NOTE]
> Como alternativa, você pode usar a [propriedade Office.context.diagnostics](#diagnostics-contextinformation) para obter a plataforma.

##### <a name="type"></a>Tipo

*   [PlatformType](/javascript/api/office/office.platformtype?view=outlook-js-1.7&preserve-view=true)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1,5|
|[Modo do Outlook aplicável](../../../outlook/outlook-add-ins-overview.md#extension-points)| Escrever ou Ler|

##### <a name="example"></a>Exemplo

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a>requirements: [RequirementsSetSupport](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.7&preserve-view=true)

Fornece um método para determinar quais conjuntos de requisitos são suportados no aplicativo e na plataforma atual.

##### <a name="type"></a>Tipo

*   [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.7&preserve-view=true)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Modo do Outlook aplicável](../../../outlook/outlook-add-ins-overview.md#extension-points)| Escrever ou Ler|

##### <a name="example"></a>Exemplo

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a>roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.7&preserve-view=true)

Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.

O objeto permite que você armazene e acesse dados para um complemento de email armazenado na caixa de correio de um usuário, de modo que está disponível para esse complemento quando ele está sendo executado de qualquer cliente Outlook usado para acessar essa caixa de `RoamingSettings` correio.

##### <a name="type"></a>Tipo

*   [RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.7&preserve-view=true)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Nível de permissão mínimo](../../../outlook/understanding-outlook-add-in-permissions.md)| Restrito|
|[Modo do Outlook aplicável](../../../outlook/outlook-add-ins-overview.md#extension-points)| Escrever ou Ler|

<br>

---
---

#### <a name="ui-ui"></a>interface do usuário: [interface do usuário](/javascript/api/office/office.ui?view=outlook-js-1.7&preserve-view=true)

Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus Office de usuário.

##### <a name="type"></a>Tipo

*   [UI](/javascript/api/office/office.ui?view=outlook-js-1.7&preserve-view=true)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Modo do Outlook aplicável](../../../outlook/outlook-add-ins-overview.md#extension-points)| Escrever ou Ler|
