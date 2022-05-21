> [!NOTE]
> No momento, as APIs de tipos de dados só estão disponíveis na visualização pública. As APIs de visualização estão sujeitas a alterações e não se destinam ao uso em um ambiente de produção. Recomendamos que você experimente apenas em ambiente de teste e desenvolvimento. Não use APIs de visualização em um ambiente de produção ou em documentos essenciais aos negócios.
>
> Para usar APIs de visualização:
>
> - Você deve fazer referência à biblioteca **beta** na rede de distribuição de conteúdo (CDN) (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). O [arquivo de definição de tipo](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) da compilação TypeScript e IntelliSense pode ser encontrado na CDN e [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts). Você pode instalar esses tipos com `npm install --save-dev @types/office-js-preview`. Para obter informações adicionais, confira o arquivo Leiame do pacote NPM [@microsoft/office-js](https://www.npmjs.com/package/@microsoft/office-js).
> - Pode ser necessário ingressar no [programa Office Insider](https://insider.office.com) para acessar builds mais recentes do Office.
>
> Para testar os tipos de dados do Office no Windows, você deve ter um número de build do Excel maior ou igual a 16.0.14626.10000. Para testar os tipos de dados do Office no Mac, você deve ter um número de build do Excel maior ou igual a 16.55.21102600.