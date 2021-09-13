---
title: Office.context - conjunto de requisitos de visualização
description: Office. Membros do objeto Context disponíveis para Outlook de usuário usando conjunto de requisitos de visualização da API de Caixa de Correio.
ms.date: 12/03/2020
ms.localizationpriority: medium
ms.openlocfilehash: f09e84062bc12f3d69adbbbedd6d02679362a728
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149059"
---
# <a name="context-mailbox-preview-requirement-set"></a>context (Conjunto de requisitos de visualização de caixa de correio)

### <a name="officecontext"></a>[Office](office.md).context

Office.context fornece interfaces compartilhadas que são usadas por complementos em todos os Office aplicativos. Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true).

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Modo do Outlook aplicável](../../../outlook/outlook-add-ins-overview.md#extension-points)| Escrever ou Ler|

## <a name="properties"></a>Propriedades

| Propriedade | Modos | Tipo de retorno | Minimum<br>conjunto de requisitos |
|---|---|---|:---:|
| [auth](#auth-auth) | Escrever<br>Ler | [Auth](/javascript/api/office/office.auth?view=outlook-js-preview&preserve-view=true) | [IdentityAPI 1.3](../../requirement-sets/identity-api-requirement-sets.md) |
| [contentLanguage](#contentlanguage-string) | Escrever<br>Ler | Cadeia de Caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [diagnostics](#diagnostics-contextinformation) | Escrever<br>Ler | [ContextInformation](/javascript/api/office/office.contextinformation?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayLanguage](#displaylanguage-string) | Escrever<br>Ler | Cadeia de Caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [host](#host-hosttype) | Escrever<br>Ler | [HostType](/javascript/api/office/office.hosttype?view=outlook-js-preview&preserve-view=true) | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [mailbox](office.context.mailbox.md) | Escrever<br>Ler | [Caixa de Correio](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [officeTheme](#officetheme-officetheme) | Escrever<br>Ler | [OfficeTheme](/javascript/api/office/office.officetheme?view=outlook-js-preview&preserve-view=true) | [Visualização](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [platform](#platform-platformtype) | Escrever<br>Ler | [PlatformType](/javascript/api/office/office.platformtype?view=outlook-js-preview&preserve-view=true) | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [requirements](#requirements-requirementsetsupport) | Escrever<br>Ler | [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [roamingSettings](#roamingsettings-roamingsettings) | Escrever<br>Ler | [RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ui](#ui-ui) | Escrever<br>Ler | [UI](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a>Detalhes da propriedade

#### <a name="auth-auth"></a>auth: [Auth](/javascript/api/office/office.auth)

Oferece suporte a [SSO (login único)](../../../outlook/authenticate-a-user-with-an-sso-token.md) fornecendo um método que permite ao aplicativo Office obter um token de acesso ao aplicativo Web do complemento. Indiretamente, isso também habilita o suplemento para acessar os dados do Microsoft Graph do usuário sem exigir que o usuário se conecte uma segunda vez.

##### <a name="type"></a>Tipo

*   [Auth](/javascript/api/office/office.auth)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| Visualização|
|[Modo do Outlook aplicável](../../../outlook/outlook-add-ins-overview.md#extension-points)| Escrever ou Ler|

##### <a name="example"></a>Exemplo

```js
Office.context.auth.getAccessTokenAsync(function(result) {
    if (result.status === "succeeded") {
        var token = result.value;
        // ...
    } else {
        console.log("Error obtaining token", result.error);
    }
});
```

<br>

---
---

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

#### <a name="diagnostics-contextinformation"></a>diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)

Obtém informações sobre o ambiente no qual o complemento está sendo executado.

##### <a name="type"></a>Tipo

*   [ContextInformation](/javascript/api/office/office.contextinformation)

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

#### <a name="host-hosttype"></a>host: [HostType](/javascript/api/office/office.hosttype)

Obtém o Office aplicativo que está hospedando o complemento.

> [!NOTE]
> Como alternativa, você pode usar a [propriedade Office.context.diagnostics](#diagnostics-contextinformation) para obter o host.

##### <a name="type"></a>Tipo

*   [HostType](/javascript/api/office/office.hosttype)

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

#### <a name="officetheme-officetheme"></a>officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)

Fornece acesso às propriedades de cores de temas do Office.

> [!NOTE]
> Esse membro só tem suporte em Outlook no Windows.

O uso de cores de tema Office permite coordenar o esquema de cores do seu add **> Office-in** com o tema atual do Office selecionado pelo usuário com a interface do usuário > Office Conta > Office, que é aplicada em todos os aplicativos cliente Office. Usar cores de temas do Office é apropriado suplementos de email e painéis de tarefas.

##### <a name="type"></a>Tipo

*   [OfficeTheme](/javascript/api/office/office.officetheme)

##### <a name="properties"></a>Propriedades

|Nome| Tipo| Descrição|
|---|---|---|
|`bodyBackgroundColor`| Cadeia de caracteres|Obtém a cor de plano de fundo do corpo de tema do Office como um tripleto hexadecimal de cores.|
|`bodyForegroundColor`| String|Obtém a cor de primeiro plano do corpo de tema do Office como um tripleto hexadecimal de cores.|
|`controlBackgroundColor`| String|Obtém a cor de plano de fundo do controle do tema do Office como um tripleto hexadecimal de cores.|
|`controlForegroundColor`| String|Obtém a cor de controle do corpo de tema do Office como um tripleto hexadecimal de cores.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| Visualização|
|[Modo do Outlook aplicável](../../../outlook/outlook-add-ins-overview.md#extension-points)| Escrever ou Ler|

##### <a name="example"></a>Exemplo

```js
function applyOfficeTheme(){
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
  var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
  var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

  // Apply body background color to a CSS class.
  $('.body').css('background-color', bodyBackgroundColor);
}
```

<br>

---
---

#### <a name="platform-platformtype"></a>plataforma: [PlatformType](/javascript/api/office/office.platformtype)

Fornece a plataforma na qual o complemento está sendo executado.

> [!NOTE]
> Como alternativa, você pode usar a [propriedade Office.context.diagnostics](#diagnostics-contextinformation) para obter a plataforma.

##### <a name="type"></a>Tipo

*   [PlatformType](/javascript/api/office/office.platformtype)

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

#### <a name="requirements-requirementsetsupport"></a>requirements: [RequirementsSetSupport](/javascript/api/office/office.requirementsetsupport)

Fornece um método para determinar quais conjuntos de requisitos são suportados no aplicativo e na plataforma atual.

##### <a name="type"></a>Tipo

*   [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)

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

#### <a name="roamingsettings-roamingsettings"></a>roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)

Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.

O objeto permite que você armazene e acesse dados para um complemento de email armazenado na caixa de correio de um usuário, de modo que está disponível para esse complemento quando ele está sendo executado de qualquer cliente Outlook usado para acessar essa caixa de `RoamingSettings` correio.

##### <a name="type"></a>Tipo

*   [RoamingSettings](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Nível de permissão mínimo](../../../outlook/understanding-outlook-add-in-permissions.md)| Restrito|
|[Modo do Outlook aplicável](../../../outlook/outlook-add-ins-overview.md#extension-points)| Escrever ou Ler|

<br>

---
---

#### <a name="ui-ui"></a>interface do usuário: [interface do usuário](/javascript/api/office/office.ui)

Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus Office de usuário.

##### <a name="type"></a>Tipo

*   [UI](/javascript/api/office/office.ui)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Modo do Outlook aplicável](../../../outlook/outlook-add-ins-overview.md#extension-points)| Escrever ou Ler|
