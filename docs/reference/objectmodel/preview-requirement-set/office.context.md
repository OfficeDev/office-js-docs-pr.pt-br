---
title: Office. Context – conjunto de requisitos de visualização
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: 5c34a7a0db5880a94ba5519059a93010a5243978
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629185"
---
# <a name="context"></a>context

### <a name="officeofficemdcontext"></a>[Office](Office.md).context

O namespace Office.context fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context).

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

##### <a name="properties"></a>Propriedades

| Propriedade | Modelos | Tipo de retorno | Mínimo<br>conjunto de requisitos |
|---|---|---|---|
| [contentLanguage](#contentlanguage-string) | Escrever<br>Ler | String | 1.0 |
| [la](#diagnostics-contextinformation) | Escrever<br>Ler | [ContextInformation](/javascript/api/office/office.contextinformation) | 1.0 |
| [displayLanguage](#displaylanguage-string) | Escrever<br>Ler | String | 1.0 |
| [principal](#host-hosttype) | Escrever<br>Ler | [HostType](/javascript/api/office/office.hosttype) | 1.0 |
| [officeTheme](#officetheme-officetheme) | Escrever<br>Ler | [OfficeTheme](/javascript/api/office/office.officetheme) | Visualização |
| [plataforma](#platform-platformtype) | Escrever<br>Ler | [PlatformType](/javascript/api/office/office.platformtype) | 1.0 |
| [atende](#requirements-requirementsetsupport) | Escrever<br>Ler | [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport) | 1.0 |
| [roamingSettings](#roamingsettings-roamingsettings) | Escrever<br>Ler | [RoamingSettings](/javascript/api/outlook/office.roamingsettings) | 1.0 |
| [ui](#ui-ui) | Escrever<br>Ler | [UI](/javascript/api/office/office.ui) | 1.0 |

### <a name="namespaces"></a>Namespaces

[auth](/javascript/api/office/office.auth): fornece suporte para [logon único (SSO)](/outlook/add-ins/authenticate-a-user-with-an-sso-token).

[Mailbox](office.context.mailbox.md): fornece acesso ao modelo de objeto do suplemento do Outlook para o Microsoft Outlook.

## <a name="property-details"></a>Detalhes da propriedade

#### <a name="contentlanguage-string"></a>contentLanguage: cadeia de caracteres

Obtém a localidade (idioma) especificada pelo usuário para edição do item.

O `contentLanguage` valor reflete a configuração de **idioma de edição** atual especificada com opções de **arquivo > > idioma** no aplicativo host do Office.

##### <a name="type"></a>Tipo

*   String

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
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
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
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
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
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
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

##### <a name="example"></a>Exemplo

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officethemejavascriptapiofficeofficeofficetheme"></a>officeTheme: [officeTheme](/javascript/api/office/office.officetheme)

Fornece acesso às propriedades de cores de temas do Office.

> [!NOTE]
> Só há suporte para esse membro no Outlook no Windows.

O uso de cores de temas do Office permite coordenar o esquema de cores do seu suplemento com o tema atual do Office selecionado pelo usuário com a **conta de arquivo > office > Office Theme UI**, que é aplicada em todos os aplicativos host do Office. Usar cores de temas do Office é apropriado suplementos de email e painéis de tarefas.

##### <a name="type"></a>Tipo

*   [OfficeTheme](/javascript/api/office/office.officetheme)

##### <a name="properties"></a>Propriedades:

|Nome| Tipo| Descrição|
|---|---|---|
|`bodyBackgroundColor`| String|Obtém a cor de plano de fundo do corpo de tema do Office como um tripleto hexadecimal de cores.|
|`bodyForegroundColor`| String|Obtém a cor de primeiro plano do corpo de tema do Office como um tripleto hexadecimal de cores.|
|`controlBackgroundColor`| String|Obtém a cor de plano de fundo do controle do tema do Office como um tripleto hexadecimal de cores.|
|`controlForegroundColor`| String|Obtém a cor de controle do corpo de tema do Office como um tripleto hexadecimal de cores.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| Visualização|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

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

#### <a name="platform-platformtypejavascriptapiofficeofficeplatformtype"></a>Platform: [platformtype](/javascript/api/office/office.platformtype)

Fornece a plataforma na qual o suplemento está sendo executado.

##### <a name="type"></a>Tipo

*   [PlatformType](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
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
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

##### <a name="example"></a>Exemplo

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.8")));
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
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
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
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|
