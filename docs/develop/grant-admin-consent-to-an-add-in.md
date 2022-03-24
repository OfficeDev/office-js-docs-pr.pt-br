---
title: Conceder consentimento ao administrador para o suplemento
description: Saiba como conceder o consentimento do administrador ao seu complemento.
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 85a230ffd3769b0013081067c88d65d38d43b760
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743780"
---
# <a name="grant-administrator-consent-to-the-add-in"></a>Conceder consentimento ao administrador para o suplemento

> [!NOTE]
> Este procedimento só é necessário durante a criação do suplemento. Quando o seu complemento de produção é implantado no AppSource ou no Centro de administração do Microsoft 365, os usuários confiarão individualmente nele ou um administrador consentiria com a organização na instalação.

Realize este procedimento *depois* de [registrar o complemento](../develop/register-sso-add-in-aad-v2.md).

1. Navegue até [o portal do Azure - Página registros de aplicativos](https://go.microsoft.com/fwlink/?linkid=2083908) para exibir o registro do aplicativo.

1. Entre com as ***credenciais de*** administrador no seu Microsoft 365 de adoção. Por exemplo, MeuNome@contoso.onmicrosoft.com.

1. Selecione o aplicativo com nome para **exibição $ADD-IN-NAME$**.

1. Na página **$ADD-IN-NAME$**, selecione permissões de **API** em seguida, na seção Permissões configuradas, escolha Conceder consentimento de administrador **para [nome do** locatário]. Selecione **Sim** para a confirmação exibida.

> [!NOTE]
> Recomendamos esse procedimento como uma prática prática prática se você estiver usando uma conta [Microsoft 365 desenvolvedor.](https://developer.microsoft.com/microsoft-365/dev-program) No entanto, se preferir, é possível fazer sideload de um add-in SSO em desenvolvimento e solicitar ao usuário um formulário de consentimento. Para obter mais informações, [consulte Sideload on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) [and Sideload on Office na Web](../testing/sideload-office-add-ins-for-testing.md).
