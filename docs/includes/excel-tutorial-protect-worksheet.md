Nesta etapa do tutorial, você adicionará outro botão à faixa de opções que, quando selecionado, executa uma função que você precisará definir para ativar e desativar a proteção da planilha.

> [!NOTE]
> Esta página descreve uma etapa individual do tutorial de suplemento do Excel. Se você chegou aqui por meio dos resultados de mecanismos de pesquisa ou via outro link direto, acesse a página de Introdução do [tutorial de suplemento do Excel](../tutorials/excel-tutorial.yml) para começá-lo do início.

## <a name="configure-the-manifest-to-add-a-second-ribbon-button"></a>Configure o manifesto para adicionar um segundo botão à faixa de opções

1. Abra o arquivo de manifesto **my-office-add-in-manifest.xml**.
2. Encontre o elemento `<Control>`. Esse elemento define o botão **Mostrar Painel de Tarefas** na faixa de opções **Início** que você usa para iniciar o suplemento. Vamos adicionar um segundo botão ao mesmo grupo na faixa de opções **Início**. Entre a marca de Controle final (`</Control>`) e a marca de Grupo final (`</Group>`), adicione a marcação a seguir.

    ```xml
    <Control xsi:type="Button" id="<!--TODO1: Unique (in manifest) name for button -->">
        <Label resid="<!--TODO2: Button label -->" />
        <Supertip>            
            <Title resid="<!-- TODO3: Button tool tip title -->" />
            <Description resid="<!-- TODO4: Button tool tip description -->" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Contoso.tpicon_16x16" />
            <bt:Image size="32" resid="Contoso.tpicon_32x32" />
            <bt:Image size="80" resid="Contoso.tpicon_80x80" />
        </Icon>
        <Action xsi:type="<!-- TODO5: Specify the type of action-->">
            <!-- TODO6: Identify the function.-->
        </Action>
    </Control>
    ```

3. Substitua `TODO1` por uma cadeia de caracteres que fornece ao botão uma ID exclusiva no arquivo de manifesto. Há apenas um outro botão no manifesto, portanto, isso não é difícil. Como nosso botão ativará ou desativará a proteção da planilha, use "ToggleProtection". Quando terminar, a marca de Controle de início inteira deve se parecer com o seguinte:

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
    ```

4. Os próximos três `TODO`s definem “resid”, que significa ID de recurso. Um recurso é uma cadeia de caracteres e você criará essas três cadeias de caracteres em uma etapa posterior. Por enquanto, você precisa fornecer IDs aos recursos. O rótulo do botão deve ser "Toggle Protection", mas a *ID* dessa cadeia de caracteres será "ProtectionButtonLabel", de forma que o elemento `Label` completo deve se parecer com o código a seguir:

    ```xml
    <Label resid="ProtectionButtonLabel" />
    ```

5. O elemento `SuperTip` define a dica de ferramenta do botão. O título da dica de ferramenta deve ser o mesmo que o rótulo do botão, por isso, usamos a mesma ID de recurso: "ProtectionButtonLabel". A descrição da dica de ferramenta será "Click to turn protection of the worksheet on and off". Mas o `ID` será "ProtectionButtonToolTip". Portanto, quando terminar, a marcação `SuperTip` inteira deve se parecer com o seguinte código: 

    ```xml
    <Supertip>            
        <Title resid="ProtectionButtonLabel" />
        <Description resid="ProtectionButtonToolTip" />
    </Supertip>
    ```

   > [!NOTE] 
   > Em um suplemento de produção,não é recomendável usar o mesmo ícone para dois botões diferentes; mas, para simplificar este tutorial, faremos isso. Portanto, a marcação `Icon` em nosso novo `Control` é apenas uma cópia do elemento `Icon` do `Control` existente. 

6. O elemento `Action` dentro do elemento `Control` original já está presente no manifesto, tem seu tipo definido como `ShowTaskpane`, mas nosso novo botão não abrirá um painel de tarefas, mas sim executará uma função personalizada criada em uma etapa posterior. Portanto, substitua `TODO5` por `ExecuteFunction`, que é o tipo de ação para botões que acionam funções personalizadas. A marca `Action` de início deve ser similar ao código abaixo:
 
    ```xml
    <Action xsi:type="ExecuteFunction">
    ```

7. O elemento `Action` original tem elementos filhos que especificam uma ID do painel de tarefas e uma URL da página que deve ser aberta no painel de tarefas. No entanto, um elemento `Action` do tipo `ExecuteFunction` tem um único elemento filho que nomeia a função executada pelo controle. Você criará essa função em uma etapa posterior e ela será chamada de `toggleProtection`. Então, substitua `TODO6` pela marcação a seguir:
 
    ```xml
    <FunctionName>toggleProtection</FunctionName>
    ```

    A marcação `Control` inteira deve ter a aparência a seguir:

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
        <Label resid="ProtectionButtonLabel" />
        <Supertip>            
            <Title resid="ProtectionButtonLabel" />
            <Description resid="ProtectionButtonToolTip" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Contoso.tpicon_16x16" />
            <bt:Image size="32" resid="Contoso.tpicon_32x32" />
            <bt:Image size="80" resid="Contoso.tpicon_80x80" />
        </Icon>
        <Action xsi:type="ExecuteFunction">
           <FunctionName>toggleProtection</FunctionName>
        </Action>
    </Control>
    ```

8. Role para baixo até a seção `Resources` do manifesto.

9. Adicione a seguinte marcação como filho do elemento `bt:ShortStrings`.

    ```xml
    <bt:String id="ProtectionButtonLabel" DefaultValue="Toggle Worksheet Protection" />
    ```

10. Adicione a seguinte marcação como filho do elemento `bt:LongStrings`.

    ```xml
    <bt:String id="ProtectionButtonToolTip" DefaultValue="Click to protect or unprotect the current worksheet." />
    ```

11. Não deixe de salvar o arquivo.

## <a name="create-the-function-that-protects-the-sheet"></a>Criar a função que protege a planilha

1. Abra o arquivo \function-file\function-file.js.

2. O arquivo já tem uma Expressão de Função Invocada Imediatamente (IFFE). Não é necessário ter uma lógica de inicialização personalizada, portanto, deixe a função atribuída a `Office.initialize` com um corpo vazio. (Mas não a exclua. A propriedade `Office.initialize` não pode ser nula ou indefinida.) *Fora da IIFE*, adicione o seguinte código. Observe que é possível especificar um parâmetro `args` para o método e a última linha do método chamará `args.completed`. Esse é um requisito para todos os comandos de suplemento do tipo **ExecuteFunction**. Ele sinaliza para o aplicativo host do Office que a função terminou e que a interface do usuário podem ficar responsiva novamente.

    ```javascript
    function toggleProtection(args) {
        Excel.run(function (context) {
            
            // TODO1: Queue commands to reverse the protection status of the current worksheet.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
        args.completed();
    }
    ```

3. Substitua `TODO1` pelo código a seguir. O código usa propriedade de proteção do objeto de planilha em um padrão de botão de alternância padrão. O `TODO2` será explicado na próxima seção.

    ```javascript
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // TODO2: Queue command to load the sheet's "protection.protected" property from
    //        the document and re-synchronize the document and task pane.

     if (sheet.protection.protected) {
        sheet.protection.unprotect();
    } else {
        sheet.protection.protect();
    }
    ``` 

## <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a>Adicione código para buscar propriedades do documento em objetos de script do painel de tarefas

Em todas as funções anteriores desta série de tutoriais, você colocou em fila comandos para *gravar* no documento do Office. Cada função terminou com uma chamada para o método `context.sync()`, que envia os comandos em fila para o documento a ser executado. Entretanto, o código adicionado na última etapa chama a propriedade `sheet.protection.protected` e essa é uma grande diferença das funções anteriores que você escreveu, pois o objeto `sheet` é apenas um objeto de proxy que existe no script do seu painel de tarefas. Ele não sabe qual é o estado real de proteção do documento, portanto, sua propriedade `protection.protected` não pode ter um valor real. É necessário primeiro buscar o status de proteção do documento e definir o valor de `sheet.protection.protected`. Somente então será possível chamar `sheet.protection.protected` sem causar uma exceção. Esse processo de busca tem três etapas:

   1. Coloque em fila um comando para carregar (ou seja, fetch) as propriedades que seu código precisa ler.
   2. Chame o método `sync` do objeto de contexto para enviar o comando em fila para o documento para execução e retornar as informações solicitadas.
   3. Como o método `sync` é assíncrono, certifique-se de que ele tenha sido concluído antes que o código chame as propriedades que foram buscadas.

Essas etapas devem ser concluídas sempre que seu código precisar *ler* informações do documento do Office.

1. Na função `toggleProtection`, substitua `TODO2` pelo seguinte código. Observação:
   - Todos os objetos do Excel têm um método `load`. Especifique as propriedades do objeto que você deseja ler no parâmetro como uma cadeia de caracteres de nomes delimitados por vírgulas. Nesse caso, a propriedade que você precisa ler é uma subpropriedade de `protection`. Referencie a subpropriedade quase exatamente como você faria em qualquer lugar do seu código, mas usando uma barra (“/”) em vez de um ponto (".").
   - Para garantir que a lógica de botão de alternância, `sheet.protection.protected`, não seja executada até após `sync` ser concluído e o `sheet.protection.protected` ser atribuída ao valor correto buscado no documento, ele será movido (na próxima etapa) para uma função `then` que não será executada até `sync` ser concluído. 

    ```javascript
    sheet.load('protection/protected');
    return context.sync()
        .then(
            function() {
                // TODO3: Move the queued toggle logic here.
            }
        )
        // TODO4: Move the final call of `context.sync` here and ensure that it
        //        does not run until the toggle logic has been queued.
    ``` 

2. Você não pode ter duas instruções `return` no mesmo caminho de código sem ramificações, portanto, exclua a linha final `return context.sync();` no final de `Excel.run`. Você adicionará um novo `context.sync` final em uma etapa posterior.
3. Recorte a estrutura `if ... else` na função `toggleProtection` e a cole no lugar de `TODO3`.
4. Substitua `TODO4` pelo código a seguir. Observação:
   - Passar o método `sync` para uma função `then` garante que ele não seja executado até que `sheet.protection.unprotect()` ou `sheet.protection.protect()` seja enfileirado.
   - O método `then` invoca qualquer função que é passada para ele e não é recomendável que `sync` seja chamado duas vezes, portanto, remova os “()” do fim de `context.sync`.

    ```javascript
    .then(context.sync);
    ```

   Quando terminar, a função inteira deve se parecer com o seguinte:

    ```javascript
    function toggleProtection(args) {
        Excel.run(function (context) {            
          const sheet = context.workbook.worksheets.getActiveWorksheet();          
          sheet.load('protection/protected');

          return context.sync()
              .then(
                  function() {
                    if (sheet.protection.protected) {
                        sheet.protection.unprotect();
                    } else {
                        sheet.protection.protect();
                    }
                  }
              )
              .then(context.sync);
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
        args.completed();
    }
    ```


## <a name="configure-the-script-loading-html-file"></a>Configure o arquivo HTML de carregamento de script

Abra o arquivo /function-file/function-file.html. Esse é um arquivo HTML sem IU que é chamado quando o usuário pressiona o botão **Ativar/Desativar Proteção da Planilha**. O objetivo é carregar o método JavaScript que deve ser executado quando botão é pressionado. Esse arquivo não será alterado. Basta observar que a segunda marca `<script>` carrega o functionfile.js.

   > [!NOTE]
   > O arquivo function-file.html e o arquivo function-file.js carregado são executados em um processo do IE completamente separado de painel de tarefas do suplemento. Se o function-file.js foi transcompilado no mesmo arquivo bundle.js que o arquivo app.js, o suplemento precisará carregar duas cópias do arquivo bundle.js, o que anule o propósito do agrupamento. Além disso, o arquivo function-file.js não contém qualquer JavaScript incompatível com o Internet Explorer. Por esses dois motivos, esse suplemento não transcompila o function-file.js. 

## <a name="test-the-add-in"></a>Testar o suplemento

1. Feche todos os aplicativos do Office, incluindo o Excel. 
2. Para excluir o cache do Office, exclua o conteúdo da pasta de cache. Isso é necessário para limpar totalmente a versão anterior do suplemento do host. 
    - No Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.
    - No Mac: `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.
3. Se, por algum motivo, o servidor não estiver executando, em uma janela do Git Bash ou em um prompt do sistema habilitado para Node.JS, acesse a pasta **Iniciar** do projeto e execute o comando `npm start`. Não é necessário recriar o projeto, pois o único arquivo JavaScript que você alterou não faz parte do bundle.js interno.
4. Usando a nova versão do arquivo de manifesto alterado, repita o processo de sideloading usando um dos seguintes métodos. *Você deve substituir a cópia anterior do arquivo de manifesto.*
    - Windows: [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Excel Online: [Realizar sideload dos Suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)
    - iPad e Mac: [Realizar sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)
7. Abra qualquer planilha no Excel.
8. Na Faixa de Opções, em **Página Inicial**, escolha **Ativar Proteger Planilha**. Observe que a maioria dos controles na Faixa de Opções está desabilitada (e visualmente esmaecida) conforme mostrado na captura de tela abaixo. 
9. Escolha uma célula como se quisesse alterar o conteúdo. Você receberá um erro informando que a planilha está protegida.
10. Escolha **Ativar/Desativar Proteção da Planilha** novamente e os controles serão reabilitados e você poderá alterar os valores das células.

    ![Tutorial do Excel - Faixa de Opções com a Proteção Ativada](../images/excel-tutorial-ribbon-with-protection-on.png)
