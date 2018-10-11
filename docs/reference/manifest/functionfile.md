# <a name="functionfile-element"></a>Elemento FunctionFile

Especifica o arquivo de código-fonte para operações que um suplemento expõe por meio de comandos de suplemento que executam uma função JavaScript em vez de exibir a interface do usuário. O elemento **FunctionFile** é um elemento filho de [DesktopFormFactor](desktopformfactor.md) ou [MobileFormFactor](mobileformfactor.md). O atributo **resid** do elemento **FunctionFile** está definido como o valor do atributo **id** de um elemento **Url** no elemento **Resources**, que contém a URL para um arquivo HTML que contém ou carrega todas as funções JavaScript usadas pelos botões de comando do suplemento sem interface do usuário , conforme definido pelo [elemento Control](control.md).

Veja a seguir um exemplo do elemento **FunctionFile**.

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

O JavaScript no arquivo HTML indicado pelo elemento **FunctionFile** deve chamar `Office.initialize` e definir funções nomeadas que utilizem um único parâmetro: `event`. As funções devem usar a API `item.notificationMessages`  para indicar o progresso, êxito ou falha ao usuário. Ele também deve chamar `event.completed` quando concluir a execução. O nome das funções são usados no elemento **FunctionName** para botões sem interface do usuário.

Veja a seguir um exemplo de um arquivo HTML que define uma função **trackMessage**.

```js
Office.initialize = function () {
    doAuth();
}

function trackMessage (event) {
    var buttonId = event.source.id;    
    var itemId = Office.context.mailbox.item.id;
    // save this message
    event.completed();
}
```

O código a seguir mostra como implementar a função usada por **FunctionName**.

```js
// The initialize function must be run each time a new page is loaded.
(function () {
    Office.initialize = function (reason) {
        // If you need to initialize something you can do so here.
    };
})();

// Your function must be in the global namespace.
function writeText(event) {

    // Implement your custom code here. The following code is a simple example.

    Office.context.document.setSelectedDataAsync("ExecuteFunction works. Button ID=" + event.source.id,
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === "failed") {
                // Show error message.
            }
            else {
                // Show success message.
            }
        });
    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}
```

> [!IMPORTANT]
> A chamada para **event.completed** indica que você tratou o evento com sucesso. Quando uma função é chamada várias vezes, por exemplo, com vários cliques no mesmo comando de suplemento, todos os eventos são enfileirados automaticamente. O primeiro evento é executado automaticamente, enquanto os outros permanecem na fila. Quando sua função chama **event.completed**, a próxima chamada em fila para essa função é executada. Você deve implementar **event.completed**; caso contrário, sua função não será executada.