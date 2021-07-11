---
title: Persistir o estado e as configurações do suplemento
description: Aprenda a persistir dados Office aplicativos Web de complemento em execução no ambiente sem estado de um controle do navegador.
ms.date: 03/23/2021
localization_priority: Normal
ms.openlocfilehash: a5a54a07abfeefda39d24e635773bfd808b59c25
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349767"
---
# <a name="persisting-add-in-state-and-settings"></a>Persistir o estado e as configurações do suplemento

[!include[information about the common API](../includes/alert-common-api-info.md)]

Essencialmente, os suplementos do Office são aplicativos Web em execução no ambiente sem estado de um controle de navegador. Como resultado, talvez o suplemento precise persistir dados para manter a continuidade de determinados recursos ou operações entre sessões de uso do suplemento. Por exemplo, o suplemento pode ter configurações personalizadas ou outros valores que precisa salvar e recarregar na próxima vez em que for inicializado, como o modo de exibição preferido ou o local padrão de um usuário. Para fazer isso, você pode:

- Use membros da API JavaScript Office que armazenam dados como:
  - Pares de nome/valor em um recipiente de propriedades armazenado em um local que depende do tipo de suplemento.
  - XML personalizado armazenado no documento.

- Usar técnicas fornecidas pelo controle de navegador subjacente: cookies de navegador ou armazenamento Web HTML5 ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) ou [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).
    > [!NOTE]
    > O usuário pode bloquear técnicas de armazenamento baseadas em navegador, dependendo das configurações escolhidas.

Este artigo se concentra em como usar a API javaScript Office para persistir o estado do add-in no documento atual. Se você precisar manter o estado entre documentos, como controlar as preferências do usuário em todos os documentos abertos, você precisará usar uma abordagem diferente. Por exemplo, você pode usar [o SSO](sso-in-office-add-ins.md#using-the-sso-token-as-an-identity) para obter a identidade do usuário e salvar a ID do usuário e suas configurações em um banco de dados online.

## <a name="persisting-add-in-state-and-settings-with-the-office-javascript-api"></a>Persistindo o estado e as configurações do add-in com a OFFICE JavaScript

A API javaScript Office fornece os [objetos Configurações](/javascript/api/office/office.settings), [RoamingSettings](/javascript/api/outlook/office.roamingsettings)e [CustomProperties](/javascript/api/outlook/office.customproperties) para salvar o estado do complemento em sessões, conforme descrito na tabela a seguir. Em todos os casos, os valores de configurações salvos são associados à [Id](../reference/manifest/id.md) do suplemento que os criou.

|**Object**|**Suporte a tipos de suplementos**|**Local de armazenamento**|**Office suporte a aplicativos**|
|:-----|:-----|:-----|:-----|
|[Configurações](/javascript/api/office/office.settings)|conteúdo e painel de tarefas|O documento, planilha ou apresentação com o qual o suplemento está trabalhando. As configurações de add-in de conteúdo e painel de tarefas estão disponíveis para o complemento que os criou a partir do documento em que são salvos.<br/><br/>**Importante:** não armazene senhas e outras IIP (informações de identificação pessoal) confidenciais com o objeto **Settings**. Os dados salvos não ficam visíveis para os usuários finais, mas são armazenados como parte do documento, que pode ser acessado pela leitura direta do formato de arquivo do documento. Você deve limitar o uso de PII pelo suplemento e armazenar quaisquer itens de PII necessários ao suplemento somente no servidor que hospeda o suplemento como um recurso protegido pelo usuário.|Word, Excel ou PowerPoint<br/><br/> **Observação:** os suplementos de painel de tarefas para o Project 2013 não dão suporte à API **Settings** para o armazenamento do estado ou das configurações do suplemento. No entanto, para os complementos em execução no Project (bem como outros aplicativos cliente Office) você pode usar técnicas como cookies do navegador ou armazenamento da Web. Para saber mais sobre essas técnicas, confira [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings). |
|[RoamingSettings](/javascript/api/outlook/office.roamingsettings)|Outlook|A caixa de correio Exchange servidor do usuário onde o complemento está instalado. Como essas configurações são armazenadas na caixa de correio do servidor do usuário, elas podem "circular" com o usuário e estão disponíveis para o add-in quando estão sendo executados no contexto de qualquer aplicativo cliente ou navegador do Office com suporte acessando a caixa de correio desse usuário.<br/><br/> As configurações móveis de suplementos do Outlook estão disponíveis apenas para o suplemento que os criou e somente por meio da caixa de correio em que o suplemento está instalado.|Outlook|
|[CustomProperties](/javascript/api/outlook/office.customproperties)|Outlook|A mensagem, o compromisso ou o item de solicitação de reunião com o qual o suplemento está trabalhando. As propriedades personalizadas de itens de suplementos do Outlook estão disponíveis apenas para o suplemento que as criou e apenas por meio do item em que estão salvas.|Outlook|
|[CustomXmlParts](/javascript/api/office/office.customxmlparts)|painel de tarefas|O documento, planilha ou apresentação com o qual o suplemento está trabalhando. As configurações de suplementos do painel de tarefas estão disponíveis para o suplemento que as criou por meio do documento em que são salvos.<br/><br/>**Importante:** não armazene senhas e outras IIP (informações de identificação pessoal) confidenciais em uma parte XML personalizada. objeto. Os dados salvos não ficam visíveis para os usuários finais, mas são armazenados como parte do documento, que pode ser acessado pela leitura direta do formato de arquivo do documento. Você deve limitar o uso de PII pelo suplemento e armazenar quaisquer itens de PII necessários ao suplemento somente no servidor que hospeda o suplemento como um recurso protegido pelo usuário.|Word (usando a Office api comum javascript) Excel (usando a API JavaScript específica do aplicativo Excel JavaScript|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a>Os dados de configurações são gerenciados na memória no tempo de execução

> [!NOTE]
> As duas seções a seguir discutem configurações no contexto da API comum de JavaScript do Office. O aplicativo específico Excel API JavaScript também fornece acesso às configurações personalizadas. As APIs do Excel e os padrões de programação são um pouco diferentes. Para saber mais, confira [SettingCollection do Excel](/javascript/api/excel/excel.settingcollection).

Internamente, os dados no pacote de propriedades acessados com o objeto , ou objetos são armazenados como um objeto JSON (Notação de Objeto JavaScript serializado) que contém pares de `Settings` `CustomProperties` `RoamingSettings` nome/valor. O nome (chave) para cada valor deve ser um , e o valor armazenado pode ser `string` um JavaScript `string` , , , ou , mas não uma `number` `date` `object` **função**.

Este exemplo da estrutura do conjunto de propriedades contém três valores de **cadeia de caracteres** definidos nomeados como `firstName`, `location` e `defaultView`.

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

Depois que o conjunto de propriedades de configurações é salvo durante a sessão anterior do suplemento, ele pode ser carregado quando o suplemento é inicializado ou a qualquer momento depois disso durante a sessão atual do suplemento. Durante a sessão, as configurações são gerenciadas inteiramente na memória usando os métodos , e do objeto que corresponde ao tipo de configuração que você está criando `get` ( Configurações `set` , `remove` **CustomProperties** ou **RoamingSettings**).

> [!IMPORTANT]
> Para persistir quaisquer adições, atualizações ou exclusões feitas durante a sessão atual do add-in para o local de armazenamento, você deve chamar o método do objeto correspondente usado para trabalhar com esse tipo de `saveAsync` configurações. Os métodos , e operam somente na cópia na `get` memória do pacote de propriedades de `set` `remove` configurações. Se o seu add-in estiver fechado sem chamar , quaisquer alterações feitas nas configurações `saveAsync` durante essa sessão serão perdidas.

## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a>Como salvar o estado e as configurações do suplemento por documento para suplementos de conteúdo e de painel de tarefas

Para persistir as configurações de estado ou personalizadas de um suplemento de conteúdo ou de painel de tarefas para Word, Excel ou PowerPoint, use o objeto [Settings](/javascript/api/office/office.settings) e seus métodos. O pacote de propriedades criado com os métodos do objeto estão disponíveis apenas para a instância do conteúdo ou do complemento do painel de tarefas que o criou e somente no documento no qual ele `Settings` é salvo.

O objeto é carregado automaticamente como parte do objeto Document e está disponível quando o painel de tarefas ou o complemento de conteúdo `Settings` é ativado. [](/javascript/api/office/office.document) Depois que `Document` o objeto é instaurou, você pode acessar o objeto com a `Settings` propriedade [settings](/javascript/api/office/office.document#settings) do `Document` objeto. Durante o tempo de vida da sessão, você pode apenas usar os métodos , e para ler, gravar ou remover as configurações persistentes e o estado do complemento da cópia na memória do pacote de `Settings.get` `Settings.set` `Settings.remove` propriedades.

Como os métodos set e remove operam apenas em relação à cópia na memória do conjunto de propriedades de configurações, para salvar configurações novas ou alteradas no documento ao qual o suplemento está associado, você deve chamar o método [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-).

### <a name="creating-or-updating-a-setting-value"></a>Criar ou atualizar um valor de configuração

O exemplo de código a seguir mostra como usar o método [Settings.set](/javascript/api/office/office.settings#set-name--value-) para criar uma configuração chamada `'themeColor'` com um valor `'green'`. O primeiro parâmetro do método set é _name_ (Id) da configuração a ser definida ou criada, que diferencia maiúsculas de minúsculas. O segundo parâmetro é o _value_ da configuração.

```js
Office.context.document.settings.set('themeColor', 'green');
```

 A configuração com o nome especificado é criada se ainda não existir, ou seu valor é atualizado se já existir. Use o `Settings.saveAsync` método para manter as configurações novas ou atualizadas no documento.

### <a name="getting-the-value-of-a-setting"></a>Obter o valor de uma configuração

O exemplo a seguir mostra como usar o método [Settings.get](/javascript/api/office/office.settings#get-name-) para obter o valor de uma configuração chamada "themeColor". O único parâmetro do `get` método é o nome sensível a _caso_ da configuração.

```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

 O `get` método retorna o valor que foi salvo anteriormente para o nome de _configuração_ que foi passado. Se a configuração não existir, o método retornará **null**.

### <a name="removing-a-setting"></a>Remover uma configuração

O exemplo a seguir mostra como usar o método [Settings.remove](/javascript/api/office/office.settings#remove-name-) para remover uma configuração com o nome "themeColor". O único parâmetro do `remove` método é o nome sensível a _caso_ da configuração.

```js
Office.context.document.settings.remove('themeColor');
```

Nada acontecerá se a configuração não existir. Use o `Settings.saveAsync` método para persistir a remoção da configuração do documento.

### <a name="saving-your-settings"></a>Salvar suas configurações

Para salvar adições, alterações ou exclusões que o suplemento fez na cópia na memória do conjunto de propriedades de configurações durante a sessão atual, você deve chamar o método [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) para armazená-lo no documento. O único parâmetro do método é retorno de chamada , que é uma função de retorno de `saveAsync` chamada com um único parâmetro. 

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

A função anônima passada para o método como o parâmetro `saveAsync` _callback_ é executada quando a operação é concluída. O _parâmetro asyncResult_ do retorno de chamada fornece acesso a um objeto que contém `AsyncResult` o status da operação. No exemplo, a função verifica a propriedade para ver se a operação de salvar foi bem-sucedida ou falhou e exibe o resultado na página `AsyncResult.status` do complemento.

## <a name="how-to-save-custom-xml-to-the-document"></a>Como salvar XML personalizado no documento

> [!NOTE]
> Esta seção discute as partes XML no contexto da API comum do JavaScript do Office com suporte no Word. O aplicativo específico Excel API JavaScript também fornece acesso às partes XML personalizadas. As APIs do Excel e os padrões de programação são um pouco diferentes. Para saber mais, confira [Excel CustomXmlPart](/javascript/api/excel/excel.customxmlpart).

Há uma opção de armazenamento adicional quando você precisa armazenar informações que excedem os limites de tamanho do documento Configurações ou que tem um caractere estruturado. Você pode manter a marcação XML personalizada em um suplemento do painel tarefas do Word (e do Excel, mas confira a observação na parte superior desta seção). No Word, use o objeto [CustomXmlPart](/javascript/api/office/office.customxmlpart) e seus métodos (novamente, consulte a observação acima do Excel). O código a seguir cria um componente XML personalizado e exibe sua ID e seu conteúdo no divs na página. Observe que deverá haver um atributo `xmlns` na cadeia de caracteres de XML.

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

Para recuperar uma parte do XML personalizado, use o método [getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-), mas a ID é um GUID gerado quando parte de XML é criada, portanto, não é possível saber ao codificar qual é a ID. Por esse motivo, ao criar uma parte de XML, é uma prática recomendada armazenar imediatamente a ID da parte de XML como uma configuração e usar uma chave fácil de lembrar. O método a seguir mostra como fazer isso. (Mas consulte seções anteriores deste artigo para obter detalhes e práticas recomendadas ao trabalhar com configurações personalizadas.)

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

## <a name="how-to-save-settings-in-an-outlook-add-in"></a>Como salvar configurações em um Outlook de dados

Para obter informações sobre como salvar configurações em um Outlook de Outlook, consulte Manage state and [settings for an Outlook add-in](../outlook/manage-state-and-settings-outlook.md).

## <a name="see-also"></a>Confira também

- [Entendendo a API de JavaScript do Office](understanding-the-javascript-api-for-office.md)
- [Suplementos do Outlook](../outlook/outlook-add-ins-overview.md)
- [Gerenciar estado e configurações para um Outlook de dados](../outlook/manage-state-and-settings-outlook.md)
- [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
