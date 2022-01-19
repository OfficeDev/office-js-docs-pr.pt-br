
Conclua as etapas a seguir para iniciar o servidor da web local e fazer o sideload do seu suplemento.

[!INCLUDE [alert use https](alert-use-https.md)]

> [!TIP]
> Se você estiver testando o seu suplemento no Mac, execute o seguinte comando antes de continuar. O servidor Web local é iniciado quando este comando é executado.
>
> ```command&nbsp;line
> npm run dev-server
> ```

- Para testar o seu suplemento no Excel, execute o seguinte comando no diretório raiz do projeto. Isso inicia o servidor Web local e abre o Excel com seu suplemento carregado.

    ```command&nbsp;line
    npm start
    ```

- Para testar o seu suplemento no Excel em um navegador, execute o seguinte comando no diretório raiz do projeto. O servidor Web local é iniciado quando este comando é executado. Substitua “{url}” pelo URL de um documento do Excel no seu OneDrive ou uma biblioteca do SharePoint para a qual você tenha permissões.

    [!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]
