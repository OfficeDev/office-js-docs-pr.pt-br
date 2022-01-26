Se o servidor da web local já estiver em execução e seu suplemento já estiver carregado no Word, prossiga para a etapa 2. Inicie o servidor Web local e realize o sideload no seu suplemento: 

- Para testar seu suplemento no Word, execute o seguinte comando no diretório raiz do seu projeto. Isso inicia o servidor Web local (caso ainda não esteja em execução) e abre o Word com o suplemento carregado.

    ```command&nbsp;line
    npm start
    ```

- Para testar o suplemento no Word na Web, execute o seguinte comando no diretório raiz do seu projeto. O servidor Web local é iniciado quando este comando é executado. Substitua "{url}" pelo URL de um documento do Word no seu OneDrive ou uma biblioteca do SharePoint para a qual você tenha permissões.

    [!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

