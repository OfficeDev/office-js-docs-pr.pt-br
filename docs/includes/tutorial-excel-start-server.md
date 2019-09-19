Se o servidor Web local já estiver em execução e se o suplemento já estiver carregado no Excel, prossiga para a etapa 2. Caso contrário, inicie o servidor Web local e Sideload seu suplemento: 

- Para testar seu suplemento no Excel, execute o seguinte comando no diretório raiz do seu projeto. Isso inicia o servidor Web local (se ele ainda não estiver sendo executado) e abre o Excel com seu suplemento carregado.

    ```command&nbsp;line
    npm start
    ```

- Para testar seu suplemento no Excel na Web, execute o seguinte comando no diretório raiz do seu projeto. Quando você executar este comando, o servidor Web local será iniciado (se ainda não estiver sendo executado).

    ```command&nbsp;line
    npm run start:web
    ```

    Para usar seu suplemento, abra um novo documento no Excel na Web e, em seguida, Sideload seu suplemento seguindo as instruções em [suplementos do Sideload Office no Office na Web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).
