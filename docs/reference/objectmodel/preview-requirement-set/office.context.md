---
title: Office.context – conjunto de requisitos de versão prévia
description: ''
ms.date: 10/11/2018
ms.openlocfilehash: 31278c2a04f21b821d2bd22fe86ed4863f6bf4c4
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433590"
---
# <a name="context"></a>context

### <a name="officeofficemdcontext"></a>[Office](Office.md).context

O namespace Office.context fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa do namespace Office.context, confira a [Referência sobre o Office.context na API compartilhada](/javascript/api/office/office.context).

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composição ou leitura|

##### <a name="members-and-methods"></a>Membros e métodos

| Membro | Tipo |
|--------|------|
| [displayLanguage](#displaylanguage-string) | Membro |
| [officeTheme](#officetheme-object) | Membro |
| [roamingSettings](#roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings) | Membro |

### <a name="namespaces"></a>Namespaces

[mailbox](office.context.mailbox.md): fornece acesso ao modelo de objeto de suplemento do Outlook para o Microsoft Outlook e o Microsoft Outlook na Web.

### <a name="members"></a>Membros

####  <a name="displaylanguage-string"></a>displayLanguage :String

Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo host do Office.

O valor `displayLanguage` reflete a configuração atual de **Display Language** especificada com **Arquivo > Opções > Idioma** no aplicativo host do Office.

##### <a name="type"></a>Tipo:

*   Cadeia de caracteres

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composição ou leitura|

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

####  <a name="officetheme-object"></a>officeTheme :Object

Fornece acesso às propriedades de cores de temas do Office.

> [!NOTE]
> Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.

Usar as cores de tema do Office possibilita coordenar o esquema de cores de seu suplemento com o tema do Office atualmente selecionado pelo usuário em **Arquivo > Conta do Office > Tema da interface de usuário do Office **, que é aplicado a todos os aplicativos host do Office. Usar cores de temas do Office é apropriado suplementos de email e painéis de tarefas.

##### <a name="type"></a>Tipo:

*   Objeto

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
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composição ou leitura|

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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a>roamingSettings :[RoamingSettings](/javascript/api/outlook/office.RoamingSettings)

Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.

O objeto `RoamingSettings` permite armazenar e acessar os dados de um suplemento de email que está armazenado na caixa de correio do usuário, para que fiquem disponíveis para esse suplemento quando ele for executado em qualquer aplicativo host de cliente usado para acessar essa caixa de correio.

##### <a name="type"></a>Tipo:

*   [RoamingSettings](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| Restrito|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redação ou leitura|