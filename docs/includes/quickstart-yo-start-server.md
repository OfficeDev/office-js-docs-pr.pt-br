1. Abra um terminal bash na raiz do projeto (**[...]/My Office Add-in**) e execute o seguinte comando para iniciar o servidor de desenvolvimento.

    ```bash
    npm start
    ```

2. Abra o Internet Explorer ou o Microsoft Edge e acesse `https://localhost:3000`. Se a página carregar sem erros de certificado, prossiga para a próxima seção neste artigo (**Experimente**). Se o seu navegador indicar que o certificado do site não é confiável, prossiga para a próxima etapa.

3. Os Suplementos Web do Office devem usar HTTPS, não HTTP, mesmo durante o desenvolvimento. Se seu navegador indicar que o certificado do site não é confiável, adicione o certificado como confiável. Veja detalhes em [Adicionar certificados autoassinados como certificados raiz confiáveis](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).

    > [!NOTE]
    > O Chrome (navegador da Web) pode continuar a indicar que o certificado do site não é confiável, mesmo depois de concluir o processo descrito em [Adição de certificados autoassinados como certificado raiz confiável](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md). Portanto, você deve usar o Internet Explorer ou o Microsoft Edge para verificar se o certificado é confiável. 

4. Depois que o navegador carregar a página do suplemento sem erros de certificado, é possível testar o suplemento.
