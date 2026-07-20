import { test, expect } from "@playwright/test";

test.describe("Génération Excel complexes e2e", () => {
  test("Mode simple - TVA 20%", async ({ page }) => {
    // Intercepter l'appel API gemini
    await page.route("**/api/gemini", async (route) => {
      await route.fulfill({
        status: 200,
        contentType: "text/plain; charset=utf-8",
        body: `Voici la formule de TVA :
\`\`\`excel
=A1*0.2
\`\`\`
Une explication rapide de la TVA.
| Ligne | Montant HT | TVA |
|---|---|---|
| Ligne 1 | 1000 | 200 |
| Ligne 2 | 500 | 100 |

✅ Vérification : syntaxe validée.
<!-- FORMULA_EN: =A1*0.2 -->`,
      });
    });

    await page.goto("/");
    const textarea = page.locator("#prompt-input");
    await textarea.fill("TVA 20% sur montant HT");
    await page.click('button[aria-label="Générer la formule Excel"]');

    // Attendre l'affichage de la réponse
    await expect(page.locator('button[title="Télécharger l\'exemple Excel (.xlsx)"]')).toBeVisible();

    // Cliquer sur le bouton Excel et intercepter le téléchargement
    const downloadPromise = page.waitForEvent("download");
    await page.click('button[title="Télécharger l\'exemple Excel (.xlsx)"]');
    const download = await downloadPromise;
    expect(download.suggestedFilename()).toContain("formule-tva-20-sur-montant-ht");
  });

  test("Mode complexe - Amortissement", async ({ page }) => {
    // Intercepter l'appel API gemini
    await page.route("**/api/gemini", async (route) => {
      await route.fulfill({
        status: 200,
        contentType: "text/plain; charset=utf-8",
        body: `Voici le tableau d'amortissement :
\`\`\`excel
=VPM($C$5/12, $C$6*12, -$C$7)
\`\`\`
Explications.
| Ligne | Mois | Intérêts | Mensualité |
|---|---|---|---|
| Ligne 1 | 1 | 500 | 1500 |
| Ligne 2 | 2 | 499 | 1500 |

✅ Vérification : validé.
<!-- FORMULA_EN: =PMT($C$5/12, $C$6*12, -$C$7) -->
<!-- TABLE_SCHEMA: {
  "type": "complex_table",
  "title": "Tableau d'amortissement de prêt",
  "parameters": [
    { "name": "Taux", "ref": "C5", "value": 0.012, "type": "percentage" },
    { "name": "Durée", "ref": "C6", "value": 12, "type": "integer" },
    { "name": "Montant", "ref": "C7", "value": 500000, "type": "currency" }
  ],
  "columns": [
    { "header": "Mois", "type": "integer", "formula": null, "formula_en": null },
    { "header": "Intérêts", "type": "currency", "formula": "=C{row}*$C$5/12", "formula_en": "=C{row}*$C$5/12" },
    { "header": "Mensualité", "type": "currency", "formula": "=VPM($C$5/12, $C$6*12, -$C$7)", "formula_en": "=PMT($C$5/12, $C$6*12, -$C$7)" }
  ],
  "data_start_row": 10,
  "sample_rows": 2
} -->`,
      });
    });

    await page.goto("/");
    const textarea = page.locator("#prompt-input");
    await textarea.fill("amortissement de pret");
    await page.click('button[aria-label="Générer la formule Excel"]');

    await expect(page.locator('button[title="Télécharger l\'exemple Excel (.xlsx)"]')).toBeVisible();

    const downloadPromise = page.waitForEvent("download");
    await page.click('button[title="Télécharger l\'exemple Excel (.xlsx)"]');
    const download = await downloadPromise;
    expect(download.suggestedFilename()).toContain("formule-amortissement-de-pret");
  });

  test("Mode complexe invalide - Fallback simple + Toast", async ({ page }) => {
    // Intercepter l'appel API avec un TABLE_SCHEMA invalide (parameters manquant alors qu'il y a des colonnes d'input)
    await page.route("**/api/gemini", async (route) => {
      await route.fulfill({
        status: 200,
        contentType: "text/plain; charset=utf-8",
        body: `Voici le tableau :
| Ligne | Mois |
|---|---|
| Ligne 1 | 1 |

✅ Vérification : ok.
<!-- FORMULA_EN: =A1 -->
<!-- TABLE_SCHEMA: {
  "type": "complex_table",
  "title": "Erreur",
  "columns": [
    { "header": "Mois", "type": "integer", "formula": null, "formula_en": null }
  ],
  "parameters": []
} -->`,
      });
    });

    await page.goto("/");
    const textarea = page.locator("#prompt-input");
    await textarea.fill("test erreur");
    await page.click('button[aria-label="Générer la formule Excel"]');

    await expect(page.locator('button[title="Télécharger l\'exemple Excel (.xlsx)"]')).toBeVisible();

    // Cliquer sur le bouton Excel et s'attendre à une erreur toast
    await page.click('button[title="Télécharger l\'exemple Excel (.xlsx)"]');

    // Vérifier l'affichage du message d'erreur
    const toast = page.locator("text=Schéma de tableau complexe invalide");
    await expect(toast).toBeVisible();
  });
});
