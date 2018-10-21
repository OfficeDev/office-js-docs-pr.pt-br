1. Abra um terminal bash na raiz do projeto (**[…]/Meu suplemento office**) e execute o seguinte comando para iniciar o servidor de desenvolvimento.

    ```bash
    npm start
    ```

2. Abra o Internet Explorer ou Microsoft Edge e acesse `https://localhost:3000`. Se a página carregar sem nenhum erro de certificado, passe para a próxima seção deste artigo (**Experimente**). Se seu navegador indicar que o certificado do site não é confiável, vá para a etapa a seguir.

3. Os Suplementos Web do Office devem usar HTTPS, não HTTP, mesmo quando você está desenvolvendo. Se o seu navegador indicar que o certificado do site não é confiável, você precisará adicioná-lo como um certificado confiável. Veja detalhes em [Adicionar certificados autoassinados como um certificado raiz confiável](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).

    > [!NOTE]
    > O Chrome (navegador da Web) pode continuar a indicar que o certificado do site não é confiável, mesmo depois de concluído o processo descrito em [Adicionar certificados autoassinados como um certificado raiz confiável](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md). Portanto, você deve usar o Internet Explorer ou Microsoft Edge para verificar se o certificado é confiável. 

4. Depois que o navegador carregar a página do suplemento sem erros de certificado, será possível testar o suplemento.
