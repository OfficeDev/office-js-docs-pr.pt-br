---
title: Exigências especiais de suplementos no iPad
description: Conheça alguns requisitos para criar um Complemento do Office executado em um iPad.
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: fdb402f4302e7e81589d586fa1ecd5b30d4e515d
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237851"
---
# <a name="special-requirements-for-add-ins-on-the-ipad"></a>Exigências especiais de suplementos no iPad

Se o seu complemento usa apenas APIs do Office compatíveis com o iPad, os clientes podem instalá-lo nos iPads. (Confira [Especificar aplicativos do Office e requisitos de API](specify-office-hosts-and-api-requirements.md) para obter mais informações.) Se o complemento for comercializado pelo *[AppSource,](https://appsource.microsoft.com)* existem algumas práticas que você deve seguir para os complementos que podem ser instalados em iPads, além das práticas [recomendadas](../concepts/add-in-development-best-practices.md)que se aplicam a todos os Complementos do Office.

A tabela a seguir lista as tarefas a executar.

> [!NOTE]
> For information about designing Outlook add-ins that look good and work well on Outlook Mobile, see [Add-ins for Outlook Mobile](../outlook/outlook-mobile-addins.md).

|Tarefas|Descrição|Recursos|
|:-----|:-----|:-----|
|Atualize seu suplemento para dar suporte ao Office.js versão 1.1.|Atualize os arquivos de JavaScript (Office.js e arquivos .js específicos do aplicativo) e o arquivo de validação de manifesto de suplemento usados no projeto do seu Suplemento do Office para a versão 1.1.|[Atualizar a versão da API e do manifesto](update-your-javascript-api-for-office-and-manifest-schema-version.md)|
|Aplicar as práticas recomendadas de design do iOS.|Integre perfeitamente a interface do usuário do seu suplemento à experiência para iOS.| Veja a observação abaixo. |
|Otimize seu suplemento para toque.|Torne sua interface do usuário responsiva a entradas de toque, além de mouse e teclado.|[Aplicar os princípios de design da UX](../concepts/add-in-development-best-practices.md#apply-ux-design-principles)|
|Torne seu suplemento gratuito.|O Office no iPad é um canal pelo qual você pode atingir mais usuários e promover seus serviços. Esses novos usuários têm potencial para se tornarem seus clientes.|[Política de certificação 1120.2](/legal/marketplace/certification-policies#11202-acquisition-pricing-and-terms)|
|Faça com que o comércio de seu complemento seja gratuito no iPad.|Quando ele estiver em execução no iPad, seu complemento deve estar livre de compras no aplicativo, ofertas de avaliação, interface do usuário com o objetivo de aumentar as vendas para uma versão não gratuita ou links para qualquer loja online onde os usuários possam comprar ou adquirir outros conteúdos, aplicativos ou complementos. Suas páginas de Política de Privacidade e Termos de Uso também devem estar livres de qualquer interface do usuário de comércio ou links do AppSource.|[Política de certificação 1100.3](/legal/marketplace/certification-policies#11003-selling-additional-features)<br><br>Seu complemento ainda pode ter comércio em outras plataformas. Para fazer isso, teste a [propriedade Office.context.commerceAllowed](/javascript/api/office/office.context#commerceallowed) e suprime todo o comércio quando `false` retornar.|
|Envie seu complemento ao AppSource.|No Partner Center,  na página Configuração do produto, marque a caixa de seleção Disponibilizar meu produto no iOS e **Android (se aplicável)** e forneça sua ID de desenvolvedor da Apple nas configurações da conta. Revise o [Contrato de Provedor de](https://go.microsoft.com/fwlink/?linkid=715691) Aplicativos para garantir que você compreendeu os termos.|[Disponibilizar suas soluções no AppSource e no Office](/office/dev/store/submit-to-appsource-via-partner-center)|

> [!NOTE]
> Seu complemento pode servir uma interface do usuário alternativa com base no dispositivo em que ele está sendo executado. Para detectar se o seu complemento está sendo executado em um iPad, você pode usar as seguintes APIs.
>
> - var isTouchEnabled = [Office.context.touchEnabled](/javascript/api/office/office.context#touchenabled)
> - var allowCommerce = [Office.context.commerceAllowed](/javascript/api/office/office.context#commerceallowed)
>
> Em um iPad, `touchEnabled` retorna `true` e retorna `commerceAllowed` `false` .
>
> Para obter informações sobre as melhores práticas de design de interface do usuário para iPad, consulte [Design para iOS.](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)

## <a name="best-practices-for-developing-office-add-ins-that-can-run-on-ipad"></a>Práticas recomendadas para o desenvolvimento de complementos do Office que podem ser executados no iPad

Aplique as seguintes práticas recomendadas para desenvolver os complementos executados no iPad.

-  **Desenvolva e depure o complemento no Windows ou Mac e o sideload para um iPad.**

    Não é possível desenvolver o complemento diretamente em um iPad, mas você pode depurá-lo em um computador Com Windows ou Mac e fazer o sideload dele em um iPad para teste. Como um complemento executado no Office no iOS ou Mac dá suporte às mesmas APIs que um complemento em execução no Office no Windows, o código do seu complemento deve ser executado da mesma maneira nessas plataformas. Para obter detalhes, [confira Testar e depurar](../testing/test-debug-office-add-ins.md) Os Complementos do Office e realizar sideload de complementos do Office no iPad e [no Mac para teste.](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

-  **Especifique os requisitos da API no manifesto do seu suplemento ou com verificações da execução.**

    Quando você especificar requisitos de API no manifesto do seu complemento, o Office determinará se o aplicativo cliente do Office dá suporte a esses membros da API. Se os membros da API estão disponíveis no aplicativo, seu complemento estará disponível. Como alternativa, você pode executar uma verificação de tempo de execução para determinar se um método está disponível no aplicativo antes de usá-lo no seu complemento. Verificações de tempo de execução garantem que o seu complemento está sempre disponível no aplicativo e fornece funcionalidade adicional se os métodos estão disponíveis. Para saber mais, confira Especificar [aplicativos do Office e requisitos de API.](specify-office-hosts-and-api-requirements.md)
