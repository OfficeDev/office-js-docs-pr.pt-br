---
title: Requisitos especiais para suplementos no iPad
description: Saiba mais sobre os requisitos para a criação de um suplemento do Office que é executado em um iPad.
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: 25ac5767db3301352e1921411af833957c4644d0
ms.sourcegitcommit: 10463841a977e9b8415362a3ae91b0ae5eebbf89
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/04/2020
ms.locfileid: "47399568"
---
# <a name="special-requirements-for-add-ins-on-the-ipad"></a>Requisitos especiais para suplementos no iPad

Se o suplemento usar apenas as APIs do Office com suporte no iPad, os clientes poderão instalá-lo no iPads. (Consulte [especificar aplicativos do Office e requisitos de API](specify-office-hosts-and-api-requirements.md) para obter mais informações.) *Se o suplemento for comercializado por meio do [AppSource](https://appsource.microsoft.com)*, há algumas práticas que você deve seguir para os suplementos que podem ser instalados no iPads, além [das práticas recomendadas que se aplicam a todos os suplementos do Office](../concepts/add-in-development-best-practices.md).

A tabela a seguir lista as tarefas a serem executadas.

> [!NOTE]
> Para obter informações sobre como criar suplementos do Outlook que têm uma boa aparência e funcionam bem no Outlook Mobile, consulte [Add-ins for Outlook Mobile](../outlook/outlook-mobile-addins.md).

|Tarefas|Descrição|Recursos|
|:-----|:-----|:-----|
|Atualize seu suplemento para dar suporte ao Office.js versão 1.1.|Atualize os arquivos de JavaScript (Office.js e arquivos .js específicos do aplicativo) e o arquivo de validação de manifesto de suplemento usados no projeto do seu Suplemento do Office para a versão 1.1.|[Atualizar a versão da API e do manifesto](update-your-javascript-api-for-office-and-manifest-schema-version.md)|
|Aplicar práticas recomendadas de design do iOS.|Integre perfeitamente a interface do usuário do seu suplemento à experiência para iOS.| Confira a observação abaixo. |
|Otimize seu suplemento para toque.|Torne sua interface do usuário responsiva a entradas de toque, além de mouse e teclado.|[Aplicar os princípios de design da UX](../concepts/add-in-development-best-practices.md#apply-ux-design-principles)|
|Torne seu suplemento gratuito.|O Office no iPad é um canal pelo qual você pode atingir mais usuários e promover seus serviços. Esses novos usuários têm potencial para se tornarem seus clientes.|[Política de certificação 1120,2](/legal/marketplace/certification-policies#11202-acquisition-pricing-and-terms)|
|Torne seu suplemento comercial gratuito no iPad.|Quando está em execução no iPad, seu suplemento deve estar livre de compras no aplicativo, ofertas de avaliação, interface do usuário que visa fazer a venda para uma versão não livre ou links para qualquer loja online, onde os usuários podem comprar ou adquirir outros conteúdos, aplicativos ou suplementos. Sua política de privacidade e as páginas de termos de uso também devem ser livres de qualquer link de interface do usuário ou de AppSource do Commerce.|[Política de certificação 1100,3](/legal/marketplace/certification-policies#11003-selling-additional-features)<br><br>O suplemento ainda pode ter o Commerce em outras plataformas. Para fazer isso, teste a propriedade [Office. Context. adicionadas commerceallowed](/javascript/api/office/office.context#commerceallowed) e omita todos os comércio quando ele retornar `false` .|
|Envie seu suplemento para o AppSource.|No centro de parceria, na página **configuração do produto** , marque a caixa de seleção **tornar meu produto disponível no Ios e no Android (se aplicável)** e forneça sua ID de desenvolvedor da Apple em configurações de conta. Revise o [contrato do provedor de aplicativos](https://go.microsoft.com/fwlink/?linkid=715691) para se certificar de que você entendeu os termos.|[Disponibilizar suas soluções no AppSource e no Office](/office/dev/store/submit-to-appsource-via-partner-center)|

> [!NOTE]
> O suplemento pode atender a uma interface do usuário alternativa com base no dispositivo em que está sendo executado. Para detectar se o suplemento está sendo executado em um iPad, você pode usar as seguintes APIs.
>
> - var isTouchEnabled = [Office.context.touchEnabled](/javascript/api/office/office.context#touchenabled)
> - var allowCommerce = [Office.context.commerceAllowed](/javascript/api/office/office.context#commerceallowed)
>
> Em um iPad, `touchEnabled` retorna `true` e `commerceAllowed` retorna `false` .
>
> Para obter informações sobre as práticas recomendadas de design de interface do usuário para iPad, consulte [Designing for Ios](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/).

## <a name="best-practices-for-developing-office-add-ins-that-can-run-on-ipad"></a>Práticas recomendadas para o desenvolvimento de suplementos do Office que podem ser executados no iPad

Aplique as seguintes práticas recomendadas para o desenvolvimento de suplementos executados no iPad.

-  **Desenvolva e depure o suplemento no Windows ou Mac e Sideload-o para um iPad.**

    Não é possível desenvolver o suplemento diretamente em um iPad, mas você pode desenvolvê-lo e depurá-lo em um computador com Windows ou Mac e Sideload-lo a um iPad para fins de teste. Como um suplemento executado no Office no iOS ou Mac oferece suporte às mesmas APIs que um suplemento executado no Office no Windows, o código do seu suplemento deve ser executado da mesma maneira nessas plataformas. Para obter detalhes, consulte [testar e depurar suplementos do Office](../testing/test-debug-office-add-ins.md) e [suplementos do Office Sideload no iPad e no Mac para teste](../testing/sideload-an-office-add-in-on-ipad-and-mac.md).

-  **Especifique os requisitos da API no manifesto do seu suplemento ou com verificações da execução.**

    Quando você especificar os requisitos da API no manifesto do suplemento, o Office determinará se o aplicativo cliente do Office oferece suporte a esses membros da API. Se os membros da API estiverem disponíveis no aplicativo, o suplemento estará disponível. Como alternativa, você pode executar uma verificação de tempo de execução para determinar se um método está disponível no aplicativo antes de usá-lo no seu suplemento. As verificações de tempo de execução garantem que o suplemento esteja sempre disponível no aplicativo e fornecerá funcionalidade adicional se os métodos estiverem disponíveis. Para obter mais informações, consulte [especificar aplicativos do Office e requisitos de API](specify-office-hosts-and-api-requirements.md).
