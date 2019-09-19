<span data-ttu-id="399a4-101">Se o servidor Web local já estiver em execução e se o suplemento já estiver carregado no Word, vá para a etapa 2.</span><span class="sxs-lookup"><span data-stu-id="399a4-101">If the local web server is already running and your add-in is already loaded in Word, proceed to step 2.</span></span> <span data-ttu-id="399a4-102">Caso contrário, inicie o servidor Web local e Sideload seu suplemento:</span><span class="sxs-lookup"><span data-stu-id="399a4-102">Otherwise, start the local web server and sideload your add-in:</span></span> 

- <span data-ttu-id="399a4-103">Para testar seu suplemento no Word, execute o seguinte comando no diretório raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="399a4-103">To test your add-in in Word, run the following command in the root directory of your project.</span></span> <span data-ttu-id="399a4-104">Isso inicia o servidor Web local (se ele ainda não estiver sendo executado) e abre o Word com seu suplemento carregado.</span><span class="sxs-lookup"><span data-stu-id="399a4-104">This starts the local web server (if it's not already running) and opens Word with your add-in loaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

- <span data-ttu-id="399a4-105">Para testar seu suplemento no Word na Web, execute o seguinte comando no diretório raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="399a4-105">To test your add-in in Word on the web, run the following command in the root directory of your project.</span></span> <span data-ttu-id="399a4-106">Quando você executar este comando, o servidor Web local será iniciado (se ainda não estiver sendo executado).</span><span class="sxs-lookup"><span data-stu-id="399a4-106">When you run this command, the local web server will start (if it's not already running).</span></span>

    ```command&nbsp;line
    npm run start:web
    ```

    <span data-ttu-id="399a4-107">Para usar seu suplemento, abra um novo documento no Word na Web e, em seguida, Sideload seu suplemento seguindo as instruções em [suplementos do Sideload Office no Office na Web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="399a4-107">To use your add-in, open a new document in Word on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>
