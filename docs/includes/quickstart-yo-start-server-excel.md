
<span data-ttu-id="e8768-101">Conclua as etapas a seguir para iniciar o servidor da web local e fazer o sideload do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="e8768-101">Complete the following steps to start the local web server and sideload your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="e8768-102">Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="e8768-102">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="e8768-103">Se você for solicitado a instalar um certificado após executar um dos seguintes comandos, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.</span><span class="sxs-lookup"><span data-stu-id="e8768-103">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

> [!TIP]
> <span data-ttu-id="e8768-104">Se você estiver testando o seu suplemento no Mac, execute o seguinte comando antes de continuar.</span><span class="sxs-lookup"><span data-stu-id="e8768-104">If you're testing your add-in on Mac, run the following command before proceeding.</span></span> <span data-ttu-id="e8768-105">Quando você executa este comando, o servidor Web local iniciará.</span><span class="sxs-lookup"><span data-stu-id="e8768-105">When you run this command, the local web server will start.</span></span>
>
> ```command&nbsp;line
> npm run dev-server
> ```

- <span data-ttu-id="e8768-106">Para testar o seu suplemento no Excel, execute o seguinte comando no diretório raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="e8768-106">To test your add-in in Excel, run the following command in the root directory of your project.</span></span> <span data-ttu-id="e8768-107">Quando você executar este comando, o servidor Web local será iniciado (se ainda não estiver em execução) e o Excel será aberto com o seu suplemento carregado.</span><span class="sxs-lookup"><span data-stu-id="e8768-107">When you run this command, the local web server will start and Word will open with your add-in loaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

- <span data-ttu-id="e8768-108">Para testar o seu suplemento no Excel Online, execute o seguinte comando no diretório raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="e8768-108">To test your add-in in Excel Online, run the following command in the root directory of your project.</span></span> <span data-ttu-id="e8768-109">Quando você executar este comando, o servidor da Web local será iniciado (se ainda não estiver em execução).</span><span class="sxs-lookup"><span data-stu-id="e8768-109">When you run this command, the local web server will start.</span></span>

    ```command&nbsp;line
    npm run start:web
    ```

    <span data-ttu-id="e8768-110">Para usar o seu suplemento, abra um novo documento no Excel Online e, em seguida, realize o sideload no suplemento de acordo com as instruções em [Realizar Sideload nos Suplementos do Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online).</span><span class="sxs-lookup"><span data-stu-id="e8768-110">To use your add-in, open a new document in Word Online and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online).</span></span>

