---
title: Extensão de depuração de suplementos do Microsoft Office para o Visual Studio Code
description: Use a extensão do Visual Studio Code do Depurador do Microsoft Office Add-in para depurar seu Complemento do Office.
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 60f7e6646cc0bfa2740e3bac0cab5f603b32dd84
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237928"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a>Extensão de depuração de suplementos do Microsoft Office para o Visual Studio Code

A Extensão de Depurador de Add-in do Microsoft Office para Visual Studio Code permite que você depure seu Complemento do Office em relação ao Microsoft Edge com o tempo de execução original do WebView (EdgeHTML). Para obter instruções sobre a depuração no Microsoft Edge WebView2 (baseado no Chromium), [consulte este artigo](./debug-desktop-using-edge-chromium.md)

Esse modo de depuração é dinâmico, permitindo definir pontos de interrupção enquanto o código está em execução. Você pode ver as alterações em seu código imediatamente enquanto o depurador está anexado, tudo sem perder sua sessão de depuração. As alterações de código também persistem, para que você possa ver os resultados de várias alterações em seu código. A imagem a seguir mostra essa extensão em ação.

![Extensão do Depurador de Addin do Office depurando uma seção de Complementos do Excel](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a>Pré-requisitos

- [Visual Studio Code](https://code.visualstudio.com/) (deve ser executado como administrador)
- [Node.js (versão 10+)](https://nodejs.org/)
- Windows 10
- [Microsoft Edge](https://www.microsoft.com/edge)

Estas instruções presumem que você tenha experiência com o uso da linha de comando, compreenda o JavaScript básico e tenha criado um projeto de Complemento do Office antes de usar o gerador Yo Office. Se você ainda não fez isso antes, considere visitar um de nossos tutoriais, como este tutorial de Complemento [do Office do Excel.](../tutorials/excel-tutorial.md)

## <a name="install-and-use-the-debugger"></a>Instalar e usar o depurador

1. Se você precisar criar um projeto de complemento, [use o gerador Yo Office para criar um.](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator) Siga os prompts dentro da linha de comando para configurar seu projeto. Você pode escolher qualquer idioma ou tipo de projeto para atender às suas necessidades.

> [!NOTE]
> Se você já tiver um projeto, pule a etapa 1 e vá para a etapa 2.

2. Abra um prompt de comando como administrador.
   ![Opções do prompt de comando, incluindo "executar como administrador" no Windows 10](../images/run-as-administrator-vs-code.jpg)

3. Navegue até o diretório do projeto.

4. Execute o seguinte comando para abrir seu projeto no Visual Studio Code como administrador.

```command&nbsp;line
code .
```

Depois que o Visual Studio Code for aberto, navegue manualmente até a pasta do projeto.

> [!TIP]
> Para abrir o Visual Studio Code como administrador, selecione a opção **executar** como administrador ao abrir o Visual Studio Code depois de procurar no Windows.

5. No VS Code, selecione **CTRL + SHIFT + X** para abrir a barra extensões. Procure a extensão "Depurador de Complementos do Microsoft Office" e instale-a.

6. Na pasta .vscode do seu projeto, abra o **launch.jsno** arquivo. Adicione o seguinte código à `configurations` seção:

```JSON
{
  "type": "office-addin",
  "request": "attach",
  "name": "Attach to Office Add-ins",
  "port": 9222,
  "trace": "verbose",
  "url": "https://localhost:3000/taskpane.html?_host_Info=HOST$Win32$16.01$en-US$$$$0",
  "webRoot": "${workspaceFolder}",
  "timeout": 45000
}
```

7. Na seção do JSON que você acabou de copiar, encontre a seção "url". Nesta URL, você precisará substituir o texto HOST em maiúsculas pelo aplicativo que está hospedando o Seu Complemento do Office. Por exemplo, se o seu Complemento do Office for para Excel, seu valor de URL seria " https://localhost:3000/taskpane.html?_host_Info= <strong>Excel</strong>$Win 32$16.01$en-US$ \$ \$ \$ 0".

8. Abra o prompt de comando e verifique se você está na pasta raiz do seu projeto. Execute o comando `npm start` para iniciar o servidor dev. Quando o seu complemento for carregado no cliente do Office, abra o painel de tarefas.

9. Retorne ao Visual Studio Code e escolha **Exibir > Depurar** ou insira **CTRL + SHIFT + D** para alternar para o exibição de depuração.

10. Nas opções de Depuração, escolha **Anexar aos Complementos do Office.** Selecione **F5** ou **Depurar -> Iniciar Depuração** no menu para começar a depuração.

11. Definir um ponto de interrupção no arquivo do painel de tarefas do projeto. Você pode definir pontos de interrupção no VS Code ao passar o mouse ao lado de uma linha de código e selecionando o círculo vermelho que aparece.

![Um círculo vermelho aparece em uma linha de código no VS Code](../images/set-breakpoint.jpg)

12. Execute o seu complemento. Você verá que pontos de interrupção foram atingidos e poderá inspecionar variáveis locais.

## <a name="see-also"></a>Confira também

* [Testar e depurar Suplementos do Office](test-debug-office-add-ins.md)

* [Depurar suplementos usando as ferramentas de desenvolvedor no Windows 10](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

* [Depurar complementos no Windows usando o Microsoft Edge WebView2 (baseado no Chromium)](debug-desktop-using-edge-chromium.md)
