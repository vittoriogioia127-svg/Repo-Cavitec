/**
 * Cavitec brand kit
 * Module reutilisable pour toute generation de document Word ou PDF.
 *
 * Principe : la source de verite est le repo GitHub Cavitec.
 * Arborescence attendue : logos/ai (maitres), logos/svg (vectoriel), logos/png (raster).
 * Le module telecharge le manifeste brand.json et les logos, verifie la
 * coherence logo / fond, puis expose des briques pretes a l emploi.
 *
 * Usage :
 *   const brand = await require('./cavitec-brand-kit').init();
 *   ...garde: [ brand.bandeau("MEMOIRE TECHNIQUE"), ... ]
 *   ...header: brand.headerCourant("Mon titre de document")
 */

const fs = require("fs");
const path = require("path");
const { execSync } = require("child_process");
const {
  Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, ImageRun, VerticalAlign, AlignmentType, TabStopType
} = require("docx");

const REPO_RAW = "https://raw.githubusercontent.com/vittoriogioia127-svg/Repo-Cavitec/main";
const CACHE = "/home/claude/.cavitec-brand";

function telecharger(rel, dest) {
  execSync(`curl -sfL -o "${dest}" "${REPO_RAW}/${rel}"`, { stdio: "pipe" });
  if (!fs.existsSync(dest) || fs.statSync(dest).size < 100) {
    throw new Error(`Telechargement echoue : ${rel}`);
  }
}

/**
 * Verifie que le logo choisi contraste avec le fond.
 * Evite l erreur classique du logo creme pose sur une page blanche.
 */
function verifierContraste(nomLogo, luminanceLogo, fondHex) {
  const r = parseInt(fondHex.slice(0, 2), 16);
  const g = parseInt(fondHex.slice(2, 4), 16);
  const b = parseInt(fondHex.slice(4, 6), 16);
  const lumFond = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
  const ecart = Math.abs(luminanceLogo - lumFond);
  if (ecart < 0.35) {
    throw new Error(
      `Contraste insuffisant : logo "${nomLogo}" (luminance ${luminanceLogo.toFixed(2)}) ` +
      `sur fond #${fondHex} (luminance ${lumFond.toFixed(2)}). Choisir une autre variante.`
    );
  }
  return true;
}

// Les luminances sont desormais portees par le manifeste (champ luminance).

async function init() {
  fs.mkdirSync(CACHE, { recursive: true });

  const manifPath = path.join(CACHE, "brand.json");
  telecharger("brand.json", manifPath);
  const M = JSON.parse(fs.readFileSync(manifPath, "utf8"));

  // Le manifeste v2 range les logos en sous dossiers ai / svg / png.
  // Pour les documents Word et PDF, seul le png est utilise.
  const chemins = {};
  for (const [cle, def] of Object.entries(M.logos.fichiers)) {
    const rel = def.png;
    const dest = path.join(CACHE, path.basename(rel));
    telecharger(`logos/${rel}`, dest);
    chemins[cle] = dest;
  }

  /** Recupere une autre declinaison (svg pour le web, ai pour archive). */
  function fichier(cleLogo, format) {
    const def = M.logos.fichiers[cleLogo];
    if (!def) throw new Error(`Logo inconnu : ${cleLogo}`);
    if (!def[format]) throw new Error(`Format ${format} indisponible pour ${cleLogo}`);
    const dest = path.join(CACHE, path.basename(def[format]));
    if (!fs.existsSync(dest)) telecharger(`logos/${def[format]}`, dest);
    return dest;
  }

  const C = M.couleurs;
  const FONT = M.typographie.famille;

  /** Bandeau d en tete : fond brun, logo creme, etiquette typologie. */
  function bandeau(etiquette) {
    verifierContraste("blanc", M.logos.fichiers.blanc.luminance, C.brun);
    const g = M.gabarits.bandeau_entete_document;
    const nul = { style: BorderStyle.NONE };
    return new Table({
      columnWidths: [5400, 3960],
      width: { size: M.gabarits.page.largeur_utile_dxa, type: WidthType.DXA },
      borders: { top: nul, bottom: nul, left: nul, right: nul, insideHorizontal: nul, insideVertical: nul },
      rows: [new TableRow({
        children: [
          new TableCell({
            width: { size: 5400, type: WidthType.DXA },
            shading: { type: ShadingType.CLEAR, fill: C.brun },
            verticalAlign: VerticalAlign.CENTER,
            margins: { top: g.marges_dxa.haut, bottom: g.marges_dxa.bas, left: g.marges_dxa.gauche, right: 120 },
            children: [new Paragraph({
              children: [new ImageRun({
                data: fs.readFileSync(chemins.blanc), type: "png",
                transformation: { width: g.logo_largeur_px, height: g.logo_hauteur_px }
              })]
            })]
          }),
          new TableCell({
            width: { size: 3960, type: WidthType.DXA },
            shading: { type: ShadingType.CLEAR, fill: C.brun },
            verticalAlign: VerticalAlign.BOTTOM,
            margins: { top: g.marges_dxa.haut, bottom: 300, left: 120, right: g.marges_dxa.droite },
            children: [new Paragraph({
              alignment: AlignmentType.RIGHT,
              children: [new TextRun({
                text: etiquette.toUpperCase().split("").join(" "),
                font: FONT, size: g.etiquette.taille_demi_points, color: C.terre_sur_brun
              })]
            })]
          })
        ]
      })]
    });
  }

  /** Header courant : logo mono brun sur fond blanc, titre a droite. */
  function headerCourant(titre) {
    verifierContraste("mono_brun", M.logos.fichiers.mono_brun.luminance, "FFFFFF");
    const g = M.gabarits.header_courant;
    return new Paragraph({
      tabStops: [{ type: TabStopType.RIGHT, position: M.gabarits.page.largeur_utile_dxa }],
      border: { bottom: { color: C.sable, size: 4, style: BorderStyle.SINGLE, space: 3 } },
      children: [
        new ImageRun({
          data: fs.readFileSync(chemins.mono_brun), type: "png",
          transformation: { width: g.logo_largeur_px, height: g.logo_hauteur_px }
        }),
        new TextRun({ text: "\t" + titre, font: FONT, size: 17, color: C.gris })
      ]
    });
  }

  /** Controle final : aucun tiret long ou moyen dans le texte produit. */
  function controlerTexte(texte) {
    const fautes = (texte.match(/[\u2013\u2014]/g) || []).length;
    if (fautes > 0) throw new Error(`${fautes} tiret(s) long(s) ou moyen(s) detecte(s). Corriger avant livraison.`);
    return true;
  }

  return { manifeste: M, couleurs: C, police: FONT, chemins, fichier, bandeau, headerCourant, controlerTexte, verifierContraste };
}

module.exports = { init };
