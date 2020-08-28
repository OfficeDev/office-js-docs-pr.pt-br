---
title: Desenvolver suplementos do Office para iPad
description: Obtenha uma visão geral e as práticas recomendadas para criar um suplemento do Office que é executado em um iPad.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 6738cc559cc07f747e075c17419b70558dec3c66
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292782"
---
# <a name="develop-office-add-ins-for-the-ipad"></a>Desenvolver suplementos do Office para iPad


A tabela a seguir lista as tarefas a serem realizadas para desenvolver um suplemento do Office para ser executado no Office no iPad.


|**Tarefa**|**Descrição**|**Recursos**|
|:-----|:-----|:-----|
|Atualize seu suplemento para dar suporte ao Office.js versão 1.1.|Atualize os arquivos de JavaScript (Office.js e arquivos .js específicos do aplicativo) e o arquivo de validação de manifesto de suplemento usados no projeto do seu Suplemento do Office para a versão 1.1.|[Atualizar a versão da API e do manifesto](update-your-javascript-api-for-office-and-manifest-schema-version.md)|
|Aplique as práticas recomendadas de design de interface do usuário.|Integre perfeitamente a interface do usuário do seu suplemento à experiência para iOS.|[Projetar para o iOS](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)|
|Aplique as práticas recomendadas de design de suplemento.|Verifique se o suplemento fornece um valor claro, é dedicado e tem um desempenho consistente.|[Práticas recomendadas para o desenvolvimento de suplementos do Office](../concepts/add-in-development-best-practices.md)|
|Otimize seu suplemento para toque.|Torne sua interface do usuário responsiva a entradas de toque, além de mouse e teclado.|[Aplicar os princípios de design da UX](../concepts/add-in-development-best-practices.md#apply-ux-design-principles)|
|Torne seu suplemento gratuito.|O Office no iPad é um canal pelo qual você pode atingir mais usuários e promover seus serviços. Esses novos usuários têm potencial para se tornarem seus clientes.|[Política de certificação 1120,2](/legal/marketplace/certification-policies#11202-acquisition-pricing-and-terms)|
|Torne a comercialização do seu suplemento gratuita.|Seu suplemento não deve oferecer compras no aplicativo, ofertas de avaliação, interfaces de usuários com o objetivo de maximizar as vendas nem links para lojas online onde os usuários possam comprar ou adquirir outros conteúdos, aplicativos ou suplementos. Suas páginas de Política de Privacidade e Termos de Uso também não devem ter nenhuma interface de usuário destinada ao comércio ou links para o AppSource.|[Política de certificação 1100,3](/legal/marketplace/certification-policies#11003-selling-additional-features)|
|Enviar soluções ao AppSource|No centro de parceria, na página **configuração do produto** , marque a caixa de seleção **tornar meu produto disponível no Ios e no Android (se aplicável)** e forneça sua ID de desenvolvedor da Apple em configurações de conta. Revise o [contrato do provedor de aplicativos](https://go.microsoft.com/fwlink/?linkid=715691) para se certificar de que você entendeu os termos.|[Disponibilizar suas soluções no AppSource e no Office](/office/dev/store/submit-to-appsource-via-partner-center)|

Seu suplemento pode permanecer como está para aplicativos do Office que estão sendo executados em outras plataformas. Você também pode fornecer uma interface de usuário diferente com base no navegador/dispositivo em que seu suplemento está sendo executado. Para detectar se seu suplemento está sendo executado em um iPad, você pode usar as seguintes APIs:
- var isTouchEnabled = [Office.context.touchEnabled](/javascript/api/office/office.context#touchenabled)
- var allowCommerce = [Office.context.commerceAllowed](/javascript/api/office/office.context#commerceallowed)


## <a name="best-practices-for-developing-office-add-ins-for-ios-and-mac"></a>Práticas recomendadas para desenvolver Suplementos do Office para iOS e Mac

Aplique as seguintes práticas recomendadas para desenvolver suplementos para execução no iOS:


-  **Use o Visual Studio para desenvolver seu suplemento.**

    Se você desenvolver seu suplemento com o Visual Studio, você pode [definir pontos de interrupção e depurar seu código](../develop/debug-office-add-ins-in-visual-studio.md) em um aplicativo cliente do Office em execução no Windows, antes de Sideload seu suplemento no iPad ou Mac. Como um suplemento executado no Office no iOS ou Mac oferece suporte às mesmas APIs que um suplemento executado no Office no Windows, o código do seu suplemento deve ser executado da mesma maneira em ambas as plataformas.

-  **Especifique os requisitos da API no manifesto do seu suplemento ou com verificações da execução.**

    Quando você especificar os requisitos da API no manifesto do suplemento, o Office determinará se o aplicativo cliente do Office oferece suporte a esses membros da API. Se os membros da API estiverem disponíveis no aplicativo, o suplemento estará disponível. Como alternativa, você pode executar uma verificação de tempo de execução para determinar se um método está disponível no aplicativo antes de usá-lo no seu suplemento. As verificações de tempo de execução garantem que o suplemento esteja sempre disponível no aplicativo e fornecerá funcionalidade adicional se os métodos estiverem disponíveis. Para obter mais informações, consulte [especificar aplicativos do Office e requisitos de API](specify-office-hosts-and-api-requirements.md).

Para ter acesso às práticas recomendadas gerais de desenvolvimento de suplementos, confira [Práticas recomendadas para desenvolvimento de Suplementos do Office](../concepts/add-in-development-best-practices.md).


## <a name="see-also"></a>Confira também

- [Realizar sideload de um suplemento do Office no iPad e no Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)  
- [Depurar suplementos do Office no iPad e no Mac](../testing/debug-office-add-ins-on-ipad-and-mac.md)
