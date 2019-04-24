---
title: Realizar o sideload de suplementos do Office em um iPad ou Mac para teste
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: e2f9ee912395e0f54130f0e78109cab4479b6567
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449919"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a>Realizar o sideload de suplementos do Office em um iPad ou Mac para teste

Para ver como seu suplemento será executado no Office para iOS, você pode realizar o sideload do manifesto do seu suplemento em um iPad usando o iTunes, ou realizar o sideload do manifesto do suplemento diretamente no Office para Mac. Esta ação não permite definir pontos de interrupção e depurar o código do seu suplemento enquanto ele estiver em execução, mas é possível ver como ele se comporta e verificar se a interface do usuário é utilizável e está sendo processada adequadamente. 

## <a name="prerequisites-for-office-for-ios"></a>Pré-requisitos do Office para iOS

- Um computador com Windows ou Mac com [iTunes](https://www.apple.com/itunes/download/) instalado.
    
- Um iPad executando o iOS 8.2 ou posterior com [Excel para iPad](https://itunes.apple.com/us/app/microsoft-excel/id586683407?mt=8) instalado e um cabo de sincronização.
    
- O arquivo de manifesto .xml para o suplemento que você deseja testar.
    

## <a name="prerequisites-for-office-for-mac"></a>Pré-requisitos do Office para Mac

- Um Mac executando o OS X v10.10 "Yosemite" ou posterior com [Office para Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) instalado.
    
- Word para Mac versão 15.18 (160109).
   
- Excel para Mac versão 15.19 (160206).

- PowerPoint para Mac versão 15.24 (160614)
    
- O arquivo de manifesto .xml para o suplemento que você deseja testar.
    

## <a name="sideload-an-add-in-on-excel-or-word-for-ipad"></a>Realizar um sideload de um suplemento no Excel ou no Word para iPad

1. Use um cabo de sincronização para conectar seu iPad ao computador. Se estiver conectando o iPad ao computador pela primeira vez, será solicitado a responder **Confiar Neste Computador?** Escolha **Confiar** para continuar.

2. No iTunes, escolha o ícone do **iPad** abaixo da barra de menus.

3. Em **Ajustes** no lado esquerdo do iTunes, escolha **Aplicativos**.

4. No lado direito do iTunes, role para baixo até **Compartilhamento de Arquivos**, e escolha **Excel** ou **Word** na coluna **Aplicativos**.

5. Na parte inferior da coluna Documentos do **Excel** ou do **Word**, escolha **Adicionar Arquivo** e selecione o arquivo de manifesto .xml do suplemento para o qual você deseja realizar sideload. 
    
6. Abra o aplicativo Excel ou Word em seu iPad. Se já estiver executando o aplicativo Excel ou Word, escolha o botão **Início**, feche e reinicie o aplicativo.
    
7. Abra um documento.
    
8. Escolha **Suplementos** na guia **Inserir**. O suplemento com sideload está disponível para inserção no cabeçalho **Desenvolvedor** na interface de usuário **Suplementos**.
    
    ![Inserir Suplementos no aplicativo do Excel](../images/excel-insert-add-in.png)


## <a name="sideload-an-add-in-on-office-for-mac"></a>Realizar sideload de um suplemento no Office para Mac

> [!NOTE]
> Para realizar o sideload de um suplemento do Outlook para Mac, confira [Realizar sideload de suplementos do Outlook para teste](/outlook/add-ins/sideload-outlook-add-ins-for-testing).

1. Abra o **Terminal** e navegue até uma das pastas a seguir, onde você salvará o arquivo de manifesto do suplemento. Se a pasta `wef` não existir em seu computador, crie-a.
    
    - Para o Word:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/documents/wef`    
    - Para o Excel:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/documents/wef`
    - Para o PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/documents/wef`
    
2. Abra a pasta no **Finder** usando o comando `open .` (incluindo o ponto final). Copie o arquivo de manifesto do suplemento nessa pasta.
    
    ![Pasta Wef no Office para Mac](../images/all-my-files.png)

3. Abra o Word e abra um documento. Reinicie o Word se já estiver em execução.
    
4. No Word, escolha **Inserir** > **Suplementos** > **Meus Suplementos** (menu suspenso) e escolha seu suplemento.
    
    ![Meus Suplementos no Office para Mac](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > Aplicativos em que foi feito o sideload não aparecerão na caixa de diálogo Meus Suplementos. Eles só ficam visíveis dentro do menu suspenso (pequena seta para baixo à direita de Meus Suplementos na guia **Inserir**). Os suplementos em que foi feito o sideload são exibidos na lista sob o título **Suplementos do Desenvolvedor** nesse menu. 
    
5. Verifique se o seu suplemento é exibido no Word.
    
    ![Suplemento do Office exibido no Office para Mac](../images/lorem-ipsum-wikipedia.png)
    
    > [!NOTE]
    > Os Suplementos muitas vezes são armazenados em cache no Office para Mac por questão de desempenho. Se você precisar forçar um recarregamento do seu suplemento durante o seu desenvolvimento, limpe a pasta `Users/<usr>/Library/Containers/com.Microsoft.OsfWebHost/Data/`. Se a pasta não existir, desmarque os arquivos da pasta `com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/`.

## <a name="see-also"></a>Confira também

- [Depurar suplementos do Office no iPad e no Mac](debug-office-add-ins-on-ipad-and-mac.md)
