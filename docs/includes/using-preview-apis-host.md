> [!NOTE]
> As APIs de visualização estão sujeitas a alterações e não se destinam ao uso em um ambiente de produção. Recomendamos que você experimente apenas em ambiente de teste e desenvolvimento. Não use APIs de visualização em um ambiente de produção ou em documentos essenciais aos negócios.
>
> Para usar APIs de visualização:
>
> - Você deve fazer referência à **biblioteca beta** no CDN ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) . O [arquivo de definição de](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) tipo para a compilação typeScript e IntelliSense é encontrado no CDN e [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts). Você pode instalar esses tipos com `npm install --save-dev @types/office-js-preview` .
> - Talvez seja necessário ingressar no programa [Office Insider](https://insider.office.com) para acessar builds Office mais recentes.