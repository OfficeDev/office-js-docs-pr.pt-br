<span data-ttu-id="6c778-101">A biblioteca da API JavaScript do Office pode ser acessada por meio da CDN (rede de entrega de conteúdo) do Office JS em: `https://appsforoffice.microsoft.com/lib/1/hosted/Office.js`.</span><span class="sxs-lookup"><span data-stu-id="6c778-101">The Office JavaScript API library can be accessed via the Office JS content delivery network (CDN) at: `https://appsforoffice.microsoft.com/lib/1/hosted/Office.js`.</span></span> <span data-ttu-id="6c778-102">Para usar as APIs JavaScript para Office em qualquer uma das páginas da Web do seu suplemento, você deve fazer referência à CDN em uma tag `<script>` na tag `<head>` da página.</span><span class="sxs-lookup"><span data-stu-id="6c778-102">To use Office JavaScript APIs within any of your add-in's web pages, you must reference the CDN in a `<script>` tag in the `<head>` tag of the page.</span></span>

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
</head>
```

> [!NOTE]
> <span data-ttu-id="6c778-103">Para usar APIs de visualização, faça referência à versão de visualização da biblioteca da API JavaScript do Office na CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.</span><span class="sxs-lookup"><span data-stu-id="6c778-103">To use preview APIs, reference the preview version of the Office JavaScript API library on the CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.</span></span>

<span data-ttu-id="6c778-104">Para obter mais informações sobre como acessar a biblioteca da API JavaScript do Office, incluindo como obter o IntelliSense, consulte [Fazendo referência à biblioteca da API JavaScript do Office a partir de sua CDN (rede de distribuição de conteúdo)](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span><span class="sxs-lookup"><span data-stu-id="6c778-104">For more information about accessing the Office JavaScript API library, including how to get IntelliSense, see [Referencing the Office JavaScript API library from its content delivery network (CDN)](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>