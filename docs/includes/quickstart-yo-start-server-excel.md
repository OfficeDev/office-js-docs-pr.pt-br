
Conclua as etapas a seguir para iniciar o servidor da web local e fazer o sideload do seu suplemento.

> [!NOTE]
> Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento. Se você for solicitado a instalar um certificado após executar um dos seguintes comandos, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.

> [!TIP]
> Se você estiver testando o seu suplemento no Mac, execute o seguinte comando antes de continuar. O servidor Web local é iniciado quando este comando é executado.
>
> ```command&nbsp;line
> npm run dev-server
> ```

- Para testar o seu suplemento no Excel, execute o seguinte comando no diretório raiz do projeto. Isso inicia o servidor Web local (caso ainda não esteja em execução) e abre o Excel com o suplemento carregado.

    ```command&nbsp;line
    npm start
    ```

- Para testar o seu suplemento no Excel em um navegador, execute o seguinte comando no diretório raiz do projeto. Quando você executar este comando, o servidor da Web local será iniciado (se ainda não estiver em execução).

    ```command&nbsp;line
    npm run start:web
    ```

    Para usar seu suplemento, abra uma nova pasta de trabalho no Excel na Web e, em seguida, realize sideload de seu suplemento seguindo as instruções em [Sideload suplementos do Office no Office Online.](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)

