---
ms.date: 05/16/2020
description: Teste seu suplemento do Office usando o Internet Explorer 11.
title: Testes do Internet Explorer 11
localization_priority: Normal
ms.openlocfilehash: 697c87d90df9aa70a7b20da5cd4c91d4445fb850
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/18/2020
ms.locfileid: "44275942"
---
# <a name="test-your-office-add-in-using-internet-explorer-11"></a>Testar o suplemento do Office usando o Internet Explorer 11

Dependendo das especificações do seu suplemento, você pode planejar o suporte a versões mais antigas do Windows e do Office, que precisam de testes no Internet Explorer 11. Isso geralmente é necessário como parte do envio do suplemento para o AppSource. Você pode usar a seguinte ferramenta de linha de comando para mudar de tempos de execução mais modernos usados pelos suplementos para o tempo de execução do Internet Explorer 11 para este teste.

## <a name="pre-requisites"></a>Pré-requisitos

- [Node.js](https://nodejs.org/) (a versão mais recente de [LTS](https://nodejs.org/about/releases))
- Um editor de códigos. Recomendamos o [Visual Studio Code](https://code.visualstudio.com/)
- [Fazer parte do programa Office Insider](https://insider.office.com)

Estas instruções pressupõem que você tenha configurado um projeto de gerador do Office Yo antes. Se você ainda não fez isso antes, considere ler um início rápido, como [este para suplementos do Excel](../quickstarts/excel-quickstart-jquery.md).

## <a name="using-ie11-tooling"></a>Usando a ferramenta de IE11

1. Criar um projeto de gerador do Office Yo. Não importa o tipo de projeto selecionado, esta ferramenta funcionará com todos os tipos de projeto.

> ! Observação Se você tiver um projeto existente e quiser adicionar essa ferramenta sem criar um novo projeto, pule esta etapa e vá para a próxima etapa. 

2. Na pasta raiz do seu novo projeto, execute o seguinte na linha de comando:

```command&nbsp;line
office-add-dev-settings webview manifest.xml ie
```
Você verá uma observação na linha de comando que o tipo de modo de exibição da Web agora está definido como IE.

> ! Tip Não é necessário usar essa ferramenta, mas ela deve ajudar a depurar a maioria dos problemas relacionados ao tempo de execução do Internet Explorer 11. Para uma robustez completa, você deve testar usando um computador com uma cópia do Windows 7 e do Office 2013 instalados.

## <a name="command-settings"></a>Configurações de comando

Se você tiver um caminho de manifesto diferente, especifique-o no comando, conforme mostrado a seguir:

`office-add-dev-settings webview [path to your manifest] ie`

O `office-addin-dev-settings webview` comando também pode ter vários tempos de execução como argumentos:

- i
- vertical
- Padrão.

## <a name="see-also"></a>Confira também
* [Testar e depurar Suplementos do Office](test-debug-office-add-ins.md)
* [Realizar sideload de suplementos do Office para teste](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [Depurar suplementos usando as ferramentas de desenvolvedor no Windows 10](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
* [Anexar um depurador do painel de tarefas](attach-debugger-from-task-pane.md)