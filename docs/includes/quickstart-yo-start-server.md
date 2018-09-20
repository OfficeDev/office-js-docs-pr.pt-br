1. Abra um terminal bash na raiz do projeto (**[…] /Meu suplemento office**) e execute o seguinte comando para iniciar o servidor de desenvolvimento.

    ```bash
    npm start
    ```

    Isso iniciará um servidor Web em `https://localhost:3000` e abrirá seu navegador padrão com esse endereço.

2. Os Suplementos Web do Office devem usar HTTPS, não HTTP, mesmo quando você está desenvolvendo. Se seu navegador indicar que o certificado do site não é confiável, adicione o certificado como confiável. Veja detalhes em [Adicionar certificados autoassinados como certificados raiz confiáveis](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).

    > [!NOTE]
    > O Chrome (navegador da Web) pode continuar a indicar que o certificado do site não é confiável, mesmo depois de concluir o processo descrito em [Adição de certificados autoassinados como certificado raiz confiável](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md). Você pode ignorar esse aviso no Chrome e verificar se o certificado é confiável ao navegar até `https://localhost:3000` no Microsoft Edge ou no Internet Explorer. 

3. Depois que o navegador carregar a página do suplemento sem erros de certificado, será possível testar o suplemento. 
