# Catálogo técnico de serviços

Blocos de escopo, entregáveis e normas para preencher o YAML da proposta. Use como base e ajuste
ao caso concreto — nunca copie um bloco que não corresponda ao serviço contratado.

## 1. Georreferenciamento de imóvel rural (SIGEF/INCRA)

**Etapas:** planejamento e análise documental (matrícula, CCIR, ITR, plantas) → pesquisa de
sobreposição na base SIGEF → implantação/rastreio GNSS (base e transporte de coordenadas) →
levantamento de limites com confrontantes → processamento e ajustamento (MMQ) → memorial
descritivo, planilha analítica e planta → submissão e acompanhamento no SIGEF até certificação.

**Entregáveis:** memorial descritivo georreferenciado; planilha analítica de vértices (ODS/SHP);
arquivo de vértices para o SIGEF; planta em PDF e DWG; certificado do INCRA; ART recolhida.

**Normas:** NTGIR (INCRA, edição vigente); Manual Técnico de Limites e Confrontações; Manual do
SIGEF; ABNT NBR 13133; Lei 10.267/2001; Decreto-Lei 9.760/1946 (quando imóvel lindeiro à União).

**Cuidados:** área registral ≠ área levantada (explique a diferença e o efeito na retificação);
certificação não é garantida (sobreposição/notificações); presença de confrontantes é do cliente.

## 2. Levantamento topográfico (urbano ou rural)

**Etapas:** reconhecimento e implantação de apoio (GNSS/estação total) → levantamento planialtimétrico
de limites, edificações e acidentes → cálculo e desenho → planta com quadro de áreas.

**Entregáveis:** planta planialtimétrica (PDF/DWG), lista de coordenadas, memorial descritivo
(quando solicitado), arquivo DXF/SHP.

**Normas:** ABNT NBR 13133; NBR 14166 (rede de referência cadastral); SIRGAS2000.

**Cuidados:** definir datum e sistema de projeção no documento; precisão exigida varia com a
finalidade (usucapião, projeto, locação de obra).

## 3. Aerolevantamento com drone (fotogrametria)

**Etapas:** planejamento de voo e checagem de restrições (SARPAS/SISANT — Lei 13.723 e RBAC-E 94)
→ implantação de pontos de apoio e alvos → voo → processamento (SfM) → ortomosaico, MDS/MDT e
curvas de nível → vetorização e planta.

**Entregáveis:** ortomosaico georreferenciado (GeoTIFF), nuvem de pontos, MDS/MDT, curvas de
nível, planta e relatório de voo com GSD e erro médio.

**Normas:** ABNT NBR 13133; RBAC-E 94 (ANAC); Decreto 87.068/1982 (se houver dados no exterior);
boas práticas ASPRS.

**Cuidados:** GSD e precisão declarados; não prometer cadastro sem pontos de apoio; cobertura
vegetal e sombreamento afetam o produto.

## 4. Cadastro técnico imobiliário e REURB

**Etapas:** levantamento de campo (unidades, confrontações, uso) → integração com base cartográfica
e cadastro municipal → planta de quadras/lotes e quadro de áreas → memoriais e cadastro individual.

**Entregáveis:** base cadastral (SHP/GeoPackage), planta geral, fichas cadastrais, relatório de
inconsistências, quadro de áreas.

**Normas:** Lei 13.465/2017 e Decreto 9.310/2018 (REURB); NBR 14166; legislação municipal.

**Cuidados:** separar honorários técnicos de taxas municipais; REURB exige articulação com o
município (não prometer prazos de titulação).

## 5. Retificação de área / desmembramento / remembramento / unificação

**Etapas:** análise da matrícula e do memorial → levantamento georreferenciado → peças técnicas →
minuta de requerimento → protocolo e acompanhamento no cartório (ou via SIGEF, se rural).

**Entregáveis:** memorial descritivo, planta, ART, minuta de requerimento, acompanhamento registral.

**Normas:** Lei 6.015/1973 (art. 213); CNJ Provimento 260/2013 e atualizações; NTGIR (rural).

**Cuidados:** emolumentos cartorários são sempre do cliente e variam por UF (tabela TJPE em PE);
declare que a aprovação registral depende do oficial.

## 6. Regularização de edificação (Habite-se / Aceite)

**Etapas:** levantamento *as built* → plantas (baixa, cortes, fachada, coberta, situação) →
memorial descritivo e laudo de habitabilidade (quando exigido) → ART → montagem e protocolo do
processo → acompanhamento de comunique-se/notificações até Habite-se/Aceite.

**Entregáveis:** jogo de plantas em PDF/DWG, memorial descritivo, ART, protocolo e certidão final.

**Normas:** legislação municipal de obras e código tributário municipal; NBR 13531/13532 (desenho);
ABNT NBR 15575 quando aplicável.

**Cuidados:** taxas municipais (análise, ISSQN da construção, expedição) são do cliente e devem
aparecer como informativas; CND do INSS quando exigida.

## 7. Perícia técnica e assistência técnica judicial

**Etapas:** leitura dos autos e quesitos → vistoria e levantamento → análise documental → laudo
com metodologia, respostas aos quesitos, anexos e ART.

**Entregáveis:** laudo técnico assinado, anexos gráficos, planilhas, ART.

**Normas:** CPC (arts. 156–158 e 465–480); NBR 14653 (avaliação, se houver valor); normas
profissionais do CONFEA/CREA.

**Cuidados:** deixar claro se é perícia (nomeado) ou assistência técnica (parte); honorários
distintos; prazo definido pelo juízo.

## 8. Estudo e implantação de loteamento

**Etapas:** viabilidade e diagnóstico → levantamento topográfico → anteprojeto e projeto geométrico
→ infraestrutura e drenagem (quando no escopo) → aprovação municipal → implantação e locação.

**Entregáveis:** estudo de viabilidade, projeto geométrico, memorial, quadro de áreas, plantas.

**Normas:** Lei 6.766/1979; legislação municipal de parcelamento; NBR 13133.

**Cuidados:** honorários por percentual do VGV (3% a 8%) exigem base de cálculo explicitada;
responsabilidades de projeto e de execução devem ser separadas.

## Serviços que normalmente NÃO entram no escopo (usar em `nao_inclusos`)

- Taxas, emolumentos e custas de terceiros (INCRA, cartório, prefeitura, CREA além da ART).
- Serviços advocatícios, pareceres jurídicos e pesquisas de cadeia dominial.
- Demarcação física, cercas, marcos e placas.
- Regularização ambiental (CAR, licenças) quando não contratada explicitamente.
- Obras, reforços estruturais e execução de infraestrutura.
- Deslocamentos extraordinários, diárias adicionais e novas convocações de confrontantes.
