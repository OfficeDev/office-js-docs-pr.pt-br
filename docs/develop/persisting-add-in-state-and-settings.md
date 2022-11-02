---
title: Manter o estado de suplemento e as configurações
description: Aprenda a persistir dados em aplicativos Web de suplemento do Office em execução no ambiente sem estado de um controle de navegador.
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: e2018e5ecf419744257cdceac31b8b1688fa65ff
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810005"
---
# <a name="persist-add-in-state-and-settings"></a>Manter o estado de suplemento e as configurações

[!include[information about the common API](../includes/alert-common-api-info.md)]

Office Add-ins are essentially web applications running in the stateless environment of a browser control. As a result, your add-in may need to persist data to maintain the continuity of certain operations or features across sessions of using your add-in. For example, your add-in may have custom settings or other values that it needs to save and reload the next time it's initialized, such as a user's preferred view or default location.
To do that, you can:

- Use membros da API JavaScript do Office que armazenam dados como:
  - Pares de nome/valor em um recipiente de propriedades armazenado em um local que depende do tipo de suplemento.
  - XML personalizado armazenado no documento.

- Usar técnicas fornecidas pelo controle de navegador subjacente: cookies de navegador ou armazenamento Web HTML5 ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) ou [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).
    > [!NOTE]
    > Alguns navegadores ou as configurações do navegador do usuário podem bloquear técnicas de armazenamento baseadas no navegador. Você deve testar a disponibilidade conforme documentado em [Usar a API de Armazenamento Da Web](https://developer.mozilla.org/docs/Web/API/Web_Storage_API/Using_the_Web_Storage_API).

Este artigo se concentra em como usar a API JavaScript do Office para persistir o estado de suplemento ao documento atual. Se você precisar persistir o estado entre documentos, como acompanhar as preferências do usuário em todos os documentos abertos, você precisará usar uma abordagem diferente. Por exemplo, você pode usar o [SSO](use-sso-to-get-office-signed-in-user-token.md) para obter a identidade do usuário e salvar a ID do usuário e suas configurações em um banco de dados online.

## <a name="persist-add-in-state-and-settings-with-the-office-javascript-api"></a>Manter o estado de suplemento e as configurações com a API JavaScript do Office

A API JavaScript do Office fornece os objetos [Configurações](/javascript/api/office/office.settings), [RoamingSettings](/javascript/api/outlook/office.roamingsettings) e [CustomProperties](/javascript/api/outlook/office.customproperties) para salvar o estado de suplemento em sessões, conforme descrito na tabela a seguir. Em todos os casos, os valores de configurações salvos são associados à [Id](/javascript/api/manifest/id) do suplemento que os criou.

|Objeto|Suporte a tipos de suplementos|Local de armazenamento|Suporte a aplicativos do Office|
|:-----|:-----|:-----|:-----|
|[Configurações](/javascript/api/office/office.settings)|-Conteúdo<br>- painel de tarefas|O documento, planilha ou apresentação com o qual o suplemento está trabalhando. As configurações de suplemento de conteúdo e painel de tarefas estão disponíveis para o suplemento que as criou a partir do documento em que são salvas.<br/><br/>**Important:** Don't store passwords and other sensitive personally identifiable information (PII) with the **Settings** object. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.|-Palavra<br>-Excel<br>-Powerpoint<br/><br/> **Observação:** os suplementos de painel de tarefas para o Project 2013 não dão suporte à API **Settings** para o armazenamento do estado ou das configurações do suplemento. No entanto, para suplementos em execução no Project (bem como em outros aplicativos cliente do Office), você pode usar técnicas como cookies de navegador ou armazenamento na Web. Para saber mais sobre essas técnicas, confira [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings). |
|[RoamingSettings](/javascript/api/outlook/office.roamingsettings)|Email|A caixa de correio do servidor exchange do usuário em que o suplemento está instalado. Como essas configurações são armazenadas na caixa de correio do servidor do usuário, elas podem "vagar" com o usuário e estão disponíveis para o suplemento quando estiver em execução no contexto de qualquer aplicativo cliente do Office ou navegador com suporte acessando a caixa de correio desse usuário.<br/><br/> As configurações móveis de suplementos do Outlook estão disponíveis apenas para o suplemento que os criou e somente por meio da caixa de correio em que o suplemento está instalado.|Outlook|
|[CustomProperties](/javascript/api/outlook/office.customproperties)|Email|The message, appointment, or meeting request item the add-in is working with. Outlook add-in item custom properties are available only to the add-in that created them, and only from the item where they are saved.|Outlook|
|[CustomXmlParts](/javascript/api/office/office.customxmlparts)|painel de tarefas|The document, spreadsheet, or presentation the add-in is working with. Task pane add-in settings are available to the add-in that created them from the document where they are saved.<br/><br/>**Important:** Don't store passwords and other sensitive personally identifiable information (PII) in a custom XML part. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.|- Word (usando a API Comum javascript do Office)<br>- Excel (usando a API JavaScript do Excel específica do aplicativo)|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a>Os dados de configurações são gerenciados na memória no tempo de execução

> [!NOTE]
> As duas seções a seguir discutem configurações no contexto da API comum de JavaScript do Office. A API JavaScript do Excel específica do aplicativo também fornece acesso às configurações personalizadas. As APIs do Excel e os padrões de programação são um pouco diferentes. Para saber mais, confira [SettingCollection do Excel](/javascript/api/excel/excel.settingcollection).

Internamente, os dados no saco de propriedades acessados com os `Settings`objetos , `CustomProperties`ou são `RoamingSettings` armazenados como um objeto JSON (Notação de Objeto JavaScript) serializado que contém pares de nome/valor. O nome (chave) para cada valor deve ser um `string`, e o valor armazenado pode ser um JavaScript `string`, `number`, `date`ou , mas `object`não uma **função**.

Este exemplo da estrutura do conjunto de propriedades contém três valores de **cadeia de caracteres** definidos nomeados como `firstName`, `location` e `defaultView`.

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

Depois que o conjunto de propriedades de configurações é salvo durante a sessão anterior do suplemento, ele pode ser carregado quando o suplemento é inicializado ou a qualquer momento depois disso durante a sessão atual do suplemento. Durante a sessão, as configurações são gerenciadas inteiramente na memória usando os `get`métodos , `set`e `remove` do objeto que correspondem ao tipo de configurações que você está criando (**Configurações**, **CustomProperties** ou **RoamingSettings**).

> [!IMPORTANT]
> Para persistir quaisquer adições, atualizações ou exclusões feitas durante a sessão atual do suplemento para o local de armazenamento, você deve chamar o `saveAsync` método do objeto correspondente usado para trabalhar com esse tipo de configurações. Os `get`métodos , `set`e `remove` operam apenas na cópia na memória do saco de propriedades de configurações. Se o suplemento estiver fechado sem chamar `saveAsync`, quaisquer alterações feitas nas configurações durante essa sessão serão perdidas.

## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a>Como salvar o estado e as configurações do suplemento por documento para suplementos de conteúdo e de painel de tarefas

Para persistir as configurações de estado ou personalizadas de um suplemento de conteúdo ou de painel de tarefas para Word, Excel ou PowerPoint, use o objeto [Settings](/javascript/api/office/office.settings) e seus métodos. O saco de propriedades criado com os métodos do `Settings` objeto está disponível apenas para a instância do conteúdo ou suplemento do painel de tarefas que o criou e somente do documento no qual ele é salvo.

O `Settings` objeto é carregado automaticamente como parte do objeto [Document](/javascript/api/office/office.document) e está disponível quando o painel de tarefas ou o suplemento de conteúdo é ativado. Depois que o `Document` objeto for instanciado, você poderá acessar o `Settings` objeto com a propriedade [configurações](/javascript/api/office/office.document#office-office-document-settings-member) do `Document` objeto. Durante o tempo de vida da sessão, você pode apenas usar os `Settings.get`métodos , `Settings.set`e `Settings.remove` para ler, gravar ou remover as configurações persistentes e o estado de suplemento da cópia na memória do saco de propriedades.

Como os métodos set e remove operam apenas em relação à cópia na memória do conjunto de propriedades de configurações, para salvar configurações novas ou alteradas no documento ao qual o suplemento está associado, você deve chamar o método [Settings.saveAsync](/javascript/api/office/office.settings#office-office-settings-saveasync-member(1)).

### <a name="creating-or-updating-a-setting-value"></a>Criar ou atualizar um valor de configuração

The following code example shows how to use the [Settings.set](/javascript/api/office/office.settings#office-office-settings-set-member(1)) method to create a setting called `'themeColor'` with a value `'green'`. The first parameter of the set method is the case-sensitive  _name_ (Id) of the setting to set or create. The second parameter is the _value_ of the setting.

```js
Office.context.document.settings.set('themeColor', 'green');
```

 A configuração com o nome especificado é criada se ainda não existir, ou seu valor é atualizado se já existir. Use o `Settings.saveAsync` método para persistir as configurações novas ou atualizadas no documento.

### <a name="getting-the-value-of-a-setting"></a>Obter o valor de uma configuração

O exemplo a seguir mostra como usar o método [Settings.get](/javascript/api/office/office.settings#office-office-settings-get-member(1)) para obter o valor de uma configuração chamada "themeColor". O único parâmetro do `get` método é o _nome_ sensível a casos da configuração.

```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

 O `get` método retorna o valor que foi salvo anteriormente para o _nome_ de configuração que foi passado. Se a configuração não existir, o método retornará **null**.

### <a name="removing-a-setting"></a>Remover uma configuração

O exemplo a seguir mostra como usar o método [Settings.remove](/javascript/api/office/office.settings#office-office-settings-remove-member(1)) para remover uma configuração com o nome "themeColor". O único parâmetro do `remove` método é o _nome_ sensível a casos da configuração.

```js
Office.context.document.settings.remove('themeColor');
```

Nada acontecerá se a configuração não existir. Use o `Settings.saveAsync` método para persistir a remoção da configuração do documento.

### <a name="saving-your-settings"></a>Salvar suas configurações

Para salvar adições, alterações ou exclusões que o suplemento fez na cópia na memória do conjunto de propriedades de configurações durante a sessão atual, você deve chamar o método [Settings.saveAsync](/javascript/api/office/office.settings#office-office-settings-saveasync-member(1)) para armazená-lo no documento. O único parâmetro do método é o `saveAsync` _retorno de chamada_, que é uma função de retorno de chamada com um único parâmetro.

```js
Office.context.document.settings.saveAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Settings save failed. Error: ' + asyncResult.error.message);
    } else {
        write('Settings saved.');
    }
});
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

A função anônima passada para o `saveAsync` método à medida que o parâmetro _de retorno de chamada_ é executado quando a operação é concluída. O parâmetro _asyncResult_ do retorno de chamada fornece acesso a um `AsyncResult` objeto que contém o status da operação. No exemplo, a função verifica a `AsyncResult.status` propriedade para ver se a operação de salvamento foi bem-sucedida ou falhou e exibe o resultado na página do suplemento.

## <a name="how-to-save-custom-xml-to-the-document"></a>Como salvar XML personalizado no documento

> [!NOTE]
> Esta seção discute as partes XML no contexto da API comum do JavaScript do Office com suporte no Word. A API JavaScript do Excel específica do aplicativo também fornece acesso às partes XML personalizadas. As APIs do Excel e os padrões de programação são um pouco diferentes. Para saber mais, confira [Excel CustomXmlPart](/javascript/api/excel/excel.customxmlpart).

Há uma opção de armazenamento adicional quando você precisa armazenar informações que excedam os limites de tamanho das Configurações do documento ou que tenha um caractere estruturado. Você pode manter a marcação XML personalizada em um suplemento do painel tarefas do Word (e do Excel, mas confira a observação na parte superior desta seção). No Word, use o objeto [CustomXmlPart](/javascript/api/office/office.customxmlpart) e seus métodos (novamente, consulte a observação acima do Excel). O código a seguir cria um componente XML personalizado e exibe sua ID e seu conteúdo no divs na página. Observe que deverá haver um atributo `xmlns` na cadeia de caracteres de XML.

```js
function createCustomXmlPart() {
    const xmlString = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    Office.context.document.customXmlParts.addAsync(xmlString,
        (asyncResult) => {
            $("#xml-id").text("Your new XML part's ID: " + asyncResult.value.id);
            asyncResult.value.getXmlAsync(
                (asyncResult) => {
                    $("#xml-blob").text(asyncResult.value);
                }
            );
        }
    );
}
```

Para recuperar uma parte do XML personalizado, use o método [getByIdAsync](/javascript/api/office/office.customxmlparts#office-office-customxmlparts-getbyidasync-member(1)), mas a ID é um GUID gerado quando parte de XML é criada, portanto, não é possível saber ao codificar qual é a ID. Por esse motivo, ao criar uma parte de XML, é uma prática recomendada armazenar imediatamente a ID da parte de XML como uma configuração e usar uma chave fácil de lembrar. O método a seguir mostra como fazer isso. (Mas consulte seções anteriores deste artigo para obter detalhes e práticas recomendadas ao trabalhar com configurações personalizadas.)

 ```js
function createCustomXmlPartAndStoreId() {
    const xmlString = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    Office.context.document.customXmlParts.addAsync(xmlString,
        (asyncResult) => {
            Office.context.document.settings.set('ReviewersID', asyncResult.id);
            Office.context.document.settings.saveAsync();
        }
    );
}
```

O código a seguir mostra como recuperar parte do XML obtendo primeiro a sua ID em uma configuração.

 ```js
function getReviewers() {
    const reviewersXmlId = Office.context.document.settings.get('ReviewersID');
    Office.context.document.customXmlParts.getByIdAsync(reviewersXmlId,
        (asyncResult) => {
            asyncResult.value.getXmlAsync(
                (asyncResult) => {
                    $("#xml-blob").text(asyncResult.value);
                }
            );
        }
    );
}
```

## <a name="how-to-save-settings-in-an-outlook-add-in"></a>Como salvar configurações em um suplemento do Outlook

Para obter informações sobre como salvar configurações em um suplemento do Outlook, consulte [Gerenciar estado e configurações para um suplemento do Outlook](../outlook/manage-state-and-settings-outlook.md).

## <a name="see-also"></a>Confira também

- [Entendendo a API de JavaScript do Office](understanding-the-javascript-api-for-office.md)
- [Suplementos do Outlook](../outlook/outlook-add-ins-overview.md)
- [Gerenciar o estado e as configurações de um suplemento do Outlook](../outlook/manage-state-and-settings-outlook.md)
- [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
