---
layout: LandingPage
ms.topic: landing-page
title: Documentação de Referência da API JavaScript do Office
description: Saiba mais sobre as APIs JavaScript do Office.
ms.date: 12/24/2019
localization_priority: Priority
ms.openlocfilehash: c10eeb5c89a74b28e9af44bf72b20a7ad610738b
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851548"
---
# <a name="api-reference-documentation"></a>Documentação de referência da API

Um suplemento pode usar as APIs de JavaScript do Office para interagir com objetos em aplicativos de host do Office. 

<ul>
    <li>As APIs<b>Específicas do host </b> fornecem objetos fortemente tipados que podem ser usados para interagir com objetos que são nativos de um aplicativo do Office específico.</li>
    <li>As APIs<b> Comuns</b> pode ser usada para acessar recursos como interface de usuário, caixas de diálogo e configurações de cliente, que são comuns entre vários tipos de aplicativos do Office.</li>
</ul>

Você deve usar APIs específicas do host sempre que possível e usar APIs comuns somente para cenários que não têm suporte em APIs específicas do host. Para obter informações mais detalhadas sobre esses dois modelos de API, confira <a href="../overview/office-add-ins-fundamentals.md#api-models">criação de Suplementos do Office</a>.

<h2>Referência da API</h2>

<ul class="panelContent cardsF cols cols3">
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/excel"><img src="../images/index/logo-excel.svg" alt="Excel API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>Referência da API do Excel</h3>
                        <p><a href="/javascript/api/excel"> APIs do JavaScript para criar suplementos do Excel.</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/outlook"><img src="../images/index/logo-outlook.svg" alt="Outlook API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>Referência da API do Outlook</h3>
                        <p><a href="/javascript/api/outlook"> APIs do JavaScript para criar suplementos do Outlook.</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/word"><img src="../images/index/logo-word.svg" alt="Word API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>Referência de API do Word</h3>
                        <p><a href="/javascript/api/word"> APIs do JavaScript para criar suplementos do Word.</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/powerpoint"><img src="../images/index/logo-powerpoint.svg" alt="PowerPoint API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>Referência do API do PowerPoint</h3>
                        <p><a href="/javascript/api/powerpoint"> APIs do JavaScript para criar suplementos do PowerPoint.</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/onenote"><img src="../images/index/logo-onenote.svg" alt="OneNote API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>Referência da API do OneNote</h3>
                        <p><a href="/javascript/api/onenote"> APIs do JavaScript para criar suplementos do OneNote.</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/office"><img src="../images/index-landing-page/i_code-blocks.svg" alt="reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>Referência da API comum</h3>
                        <p><a href="/javascript/api/office">APIs do JavaScript que podem ser usadas por qualquer suplemento do Office.</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
</ul>

<b>Observação</b>: atualmente, não há nenhuma API JavaScript específica do host para o Project. Você usará APIs comuns para criar suplementos de Project. Além disso, a API específica do host para o PowerPoint tem um escopo muito limitado. Você usará principalmente APIs comuns para criar suplementos do PowerPoint.

<h2>Especificações abertas da API</h2>

À medida que criamos e desenvolvemos novas APIs para suplementos do Office, nós as disponibilizamos em nossa página [Especificações abertas da API](openspec/openspec.md) a fim de obter os seus comentários. Descubra quais novos recursos estão no pipeline e forneça comentários sobre nossas especificações de design.