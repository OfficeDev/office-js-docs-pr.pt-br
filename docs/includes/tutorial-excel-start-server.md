Se o servidor da Web local já estiver em execução e seu suplemento já estiver carregado no Word, prossiga para a etapa 2. Inicie o servidor Web local e realize o sideload no seu suplemento: 

- Para testar o seu suplemento no Excel, execute o seguinte comando no diretório raiz do projeto. Isso inicia o servidor Web local (caso ainda não esteja em execução) e abre o Excel com o suplemento carregado.

    ```command&nbsp;line
    npm start
    ```

- Para testar seu suplemento no Excel na Web, execute o seguinte comando no diretório raiz do seu projeto. Quando você executar este comando, o servidor da Web local será iniciado (se ainda não estiver em execução).

    ```command&nbsp;line
    npm run start:web
    ```

    Para usar o seu suplemento, abra um novo documento no Excel na Web e em seguida realize o sideload no suplemento de acordo com as instruções em [Realizar Sideload nos Suplementos do Office no Office na Web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).
