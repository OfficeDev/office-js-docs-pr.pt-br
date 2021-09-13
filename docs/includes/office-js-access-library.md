A biblioteca da API JavaScript do Office pode ser acessada por meio da CDN (rede de entrega de conteúdo) do Office JS em: `https://appsforoffice.microsoft.com/lib/1/hosted/office.js`. Para usar as APIs JavaScript para Office em qualquer uma das páginas da Web do seu suplemento, você deve fazer referência à CDN em uma tag `<script>` na tag `<head>` da página.

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

> [!NOTE]
> Para usar APIs de visualização, faça referência à versão de visualização da biblioteca da API JavaScript do Office na CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.

Para obter mais informações sobre como acessar a biblioteca da API JavaScript do Office, incluindo como obter o IntelliSense, consulte [Fazendo referência à biblioteca da API JavaScript do Office a partir de sua CDN (rede de distribuição de conteúdo)](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).