const pptxgen = require("pptxgenjs");
const fs = require("fs");
const FIG = "/home/claude/tiny-nlp-research/results/figures/";
const NAVY="14213D", TEAL="1B9E77", ORANGE="D95F02", GREY="7F7F7F", LIGHT="F4F6F8", INK="1F2937", MUTED="5B6472", BLUE="1F78B4", PURPLE="7570B3", BROWN="8C564B", CYAN="17BECF", OLIVE="BCBD22";
const F="Calibri";
const pres = new pptxgen(); pres.layout = "LAYOUT_16x9"; pres.title = "Same Words, Different Bets";
const NOTES = [];

function base(title, sub, tag) {
  const s = pres.addSlide(); s.background = { color: "FFFFFF" };
  s.addShape(pres.ShapeType.rect, { x:0, y:0, w:0.12, h:5.625, fill:{color: tag ? ORANGE : TEAL}, line:{color: tag ? ORANGE : TEAL} });
  s.addText(title, { x:0.45, y:0.25, w:8.6, h:0.6, fontFace:F, fontSize:25, bold:true, color:NAVY, margin:0, isTextBox:true });
  if (sub) s.addText(sub, { x:0.45, y:0.83, w:9.1, h:0.35, fontFace:F, fontSize:12.5, color:MUTED, margin:0, isTextBox:true });
  if (tag) s.addText("BACKUP", { x:8.6, y:0.28, w:1.0, h:0.3, fontFace:F, fontSize:11, bold:true, color:ORANGE, align:"right", margin:0, isTextBox:true });
  return s;
}
function bullets(s, items, o) {
  s.addText(items.map((t,i)=>({ text:t, options:{ bullet:true, breakLine:i<items.length-1, paraSpaceAfter:7 } })),
    Object.assign({ x:0.45, y:1.35, w:9, h:3.8, fontFace:F, fontSize:15, color:INK, valign:"top", margin:0, isTextBox:true }, o||{}));
}
function stat(s, x, y, w, big, small, color) {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h:1.3, fill:{color:LIGHT}, line:{color:LIGHT}, rectRadius:0.06 });
  s.addText(big, { x:x+0.15, y:y+0.08, w:w-0.3, h:0.68, fontFace:F, fontSize:27, bold:true, color:color||TEAL, margin:0, isTextBox:true });
  s.addText(small, { x:x+0.15, y:y+0.76, w:w-0.3, h:0.5, fontFace:F, fontSize:11, color:MUTED, margin:0, valign:"top", isTextBox:true });
}
function note(s, label, timing, script) { s.addNotes(`[${label} | ${timing}]\n\n${script}`); NOTES.push([label, timing, script]); }
const chartBase = { catAxisLabelColor:MUTED, valAxisLabelColor:MUTED, catAxisLabelFontFace:F, valAxisLabelFontFace:F,
  catAxisLabelFontSize:10.5, valAxisLabelFontSize:10.5, valGridLine:{color:"E5E7EB", size:0.5}, catGridLine:{style:"none"},
  showLegend:true, legendPos:"b", legendFontFace:F, legendFontSize:10, legendColor:INK, showTitle:false };
function box(s, x, y, w, h, txt, fill, color) {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, fill:{color:fill||LIGHT}, line:{color:fill||LIGHT}, rectRadius:0.08 });
  s.addText(txt, { x, y, w, h, fontFace:F, fontSize:11.5, color:color||INK, align:"center", valign:"middle", margin:4, isTextBox:true });
}

// ================================================================ 1 Title
{ const s = pres.addSlide(); s.background = { color: NAVY };
  s.addText("Same Words, Different Bets", { x:0.7, y:1.05, w:8.6, h:1.1, fontFace:F, fontSize:40, bold:true, color:"FFFFFF", margin:0, isTextBox:true });
  s.addText("Lexicons, static embeddings and contextual pre-training for sentiment classification under scarce labels and ambiguity", { x:0.7, y:2.25, w:8.6, h:0.9, fontFace:F, fontSize:16.5, color:"C7D2E0", margin:0, isTextBox:true });
  s.addShape(pres.ShapeType.rect, { x:0.7, y:3.35, w:1.2, h:0.06, fill:{color:TEAL}, line:{color:TEAL} });
  s.addText("Given 1.2M unlabeled words: ignore them, distil them, or pre-train on them?", { x:0.7, y:3.55, w:8.6, h:0.4, fontFace:F, fontSize:14, italic:true, color:"9FB3C8", margin:0, isTextBox:true });
  s.addText("SST-2 + IMDb  |  CPU only  |  3 seeds  |  Anurag Prasad", { x:0.7, y:4.5, w:8.6, h:0.4, fontFace:F, fontSize:13, color:"9FB3C8", margin:0, isTextBox:true });
  note(s, "Slide 1 - Title", "0:00-0:45",
`Good morning. Given the same small amount of unlabeled text, what's the best bet - ignore it, turn it into a lexicon, distil it into static embeddings, or spend it on contextual pre-training?
I built a CPU-only pipeline to answer this for sentiment classification, comparing a rule-based lexicon, TF-IDF, static word embeddings trained on the same text, a tiny BERT with and without pre-training, and public checkpoints.
Beyond accuracy, I also look at robustness and - the twist - where models succeed or fail under ambiguity. Ten minutes, then five for questions.`); }

// ================================================================ 2 Problem
{ const s = base("The problem: seven questions, one protocol", "Same data, same head, same evaluation for every model");
  const rq = [["RQ1-2","Does pre-training help? How robust is it?"],["RQ3-4","How much pre-training is enough, and what does it cost?"],
              ["RQ5","How far do public checkpoints move the picture?"],["RQ6","Given the same text, what's the best family to use it with?"],
              ["RQ7","Where do errors concentrate under ambiguity, and do models know when they're unsure?"]];
  rq.forEach((r,i)=>{ const y=1.35+i*0.78;
    s.addShape(pres.ShapeType.roundRect,{x:0.45,y,w:9.1,h:0.68,fill:{color:LIGHT},line:{color:LIGHT},rectRadius:0.05});
    s.addText(r[0],{x:0.6,y,w:1.3,h:0.68,fontFace:F,fontSize:15,bold:true,color:TEAL,valign:"middle",margin:0,isTextBox:true});
    s.addText(r[1],{x:2.0,y,w:7.4,h:0.68,fontFace:F,fontSize:13.5,color:INK,valign:"middle",margin:0,isTextBox:true}); });
  note(s, "Slide 2 - Problem", "0:45-1:45",
`Seven questions under one protocol. RQ1 and 2: does masked-language-model pre-training beat random init and TF-IDF, and how robust is it? RQ3 and 4: how does the benefit change with pre-training length, and what does it cost in size and speed? RQ5: what happens with real public checkpoints? RQ6, the new one: given the same unlabeled text, is contextual pre-training even the best way to use it, compared to a lexicon or static embeddings? RQ7, also new: where do errors concentrate on ambiguous sentences, and do models know when they're unsure?
Same data subsets, same classifier head, same evaluation sets for every single model - that's what makes the comparison fair.`); }

// ================================================================ 3 Approach
{ const s = base("Approach: six families, one protocol", "Every model - lexicon to transformer - fine-tuned/fit and evaluated identically");
  const fam = [["VADER lexicon","0 task labels"],["TF-IDF word/char","+ LogReg/NB/SVM"],["word2vec / fastText","same 1.2M words"],
               ["Tiny-BERT","random init"],["Tiny-BERT","+MLM pre-train"],["BERT-Tiny /\nDistilBERT","public"]];
  fam.forEach((f,i)=>{ const x=0.45+i*1.545; box(s,x,1.35,1.42,1.05,`${f[0]}\n${f[1]}`,i>=3?"E3EEF8":LIGHT,INK); });
  s.addText("Evaluated on: clean SST-2, typos (10/30%), word-drop (20/40%), out-of-domain IMDb, and - new - ambiguity strata from fine-grained SST-5 labels (strong vs weak polarity), confidence calibration, and linguistic markers (contrast, negation).",
    { x:0.45,y:2.65,w:9.1,h:1.0,fontFace:F,fontSize:13.5,color:MUTED,margin:0,isTextBox:true, valign:"top" });
  s.addText("Tiny-BERT: 2 layers, hidden 128, 1,173,248 params. MLM pre-training: 5,000 steps, 13 min, on 1.2M unlabeled IMDb words. All decisions (regularisation, embedding choices) tuned on dev only.",
    { x:0.45,y:3.85,w:9.1,h:1.2,fontFace:F,fontSize:13,italic:true,color:MUTED,margin:0,isTextBox:true, valign:"top" });
  note(s, "Slide 3 - Approach", "1:45-2:45",
`Six families, one protocol. A VADER sentiment lexicon that never sees a training label. TF-IDF with word or character n-grams. word2vec and fastText - static embeddings trained on the exact same 1.2 million unlabeled words as our MLM pre-training. Our tiny BERT, with and without that pre-training. And public BERT-Tiny and DistilBERT.
Everything is evaluated on clean data, typos, word-drop, out-of-domain IMDb reviews, and - new for this version - ambiguity strata built from fine-grained SST-5 labels, confidence calibration, and simple linguistic markers like contrast words.
The tiny BERT has 1.17 million parameters and pre-trains in 13 minutes on one CPU.`); }

// ================================================================ 4 Plateau
{ const s = base("Pre-training starts with a long plateau", "Held-out masked-LM loss and accuracy over 5,000 steps");
  s.addImage({ path:FIG+"fig_pretrain.png", x:0.45, y:1.3, w:5.7, h:2.1, sizing:{type:"contain",w:5.7,h:2.1} });
  stat(s, 6.35, 1.4, 3.2, "ppl 77", "final held-out perplexity (from ~3,660)");
  stat(s, 6.35, 2.85, 3.2, "29%", "masked-token accuracy", ORANGE);
  bullets(s, ["First ~1,500 steps: loss stuck near the unigram entropy", "Only then does attention start to use context"], { x:0.45, y:3.55, w:6.0, h:1.4, fontSize:13 });
  note(s, "Slide 4 - Plateau", "2:45-3:15",
`Quick sanity check before results. For the first 1,500 of 5,000 pre-training steps, the loss sits near the unigram entropy - the model is only learning word frequencies. It breaks out after that, ending at perplexity 77 and 29 percent masked-token accuracy. So pre-training is doing something real; the question is whether it helps.`); }

// ================================================================ 5 Verdict
{ const s = base("Pre-training did not help; TF-IDF wins", "SST-2 test accuracy (%), mean of 3 seeds");
  const labels = ["100","300","1000","3000","6920"];
  s.addChart(pres.charts.LINE, [
    {name:"TF-IDF + LogReg", labels, values:[57.9,63.2,70.5,77.5,81.3]},
    {name:"Tiny-BERT random init", labels, values:[57.0,63.2,68.2,75.2,78.5]},
    {name:"Tiny-BERT + MLM", labels, values:[56.8,58.7,65.7,72.9,77.5]}],
    Object.assign({}, chartBase, { x:0.35, y:1.3, w:5.9, h:3.6, chartColors:[GREY,ORANGE,TEAL], lineSize:2.5, lineDataSymbolSize:7,
      valAxisMinVal:50, valAxisMaxVal:85, showCatAxisTitle:true, catAxisTitle:"labeled training sentences", catAxisTitleColor:MUTED, catAxisTitleFontSize:10 }));
  bullets(s, ["MLM vs random init: -0.3 to -4.5 pts at every budget (CIs exclude 0 at n=300,1k,3k)",
              "TF-IDF beats the pre-trained model by 3.9-4.8 pts from n=300 on"], { x:6.45, y:1.45, w:3.1, h:3.5, fontSize:13 });
  note(s, "Slide 5 - Verdict", "3:15-4:00",
`Here's the headline. From 300 labels on, the pre-trained model falls behind its own untrained twin - by 0.3 to 4.5 points, and the confidence intervals exclude zero at three of five budgets. TF-IDF then overtakes both. At laptop scale, this pre-training recipe did not pay off in the usual sense. But that's not the whole story - the next two slides are why.`); }

// ================================================================ 6 Families
{ const s = base("RQ6: the best bet depends on the label budget", "SST-2 accuracy by model family (mean of 3 seeds)");
  s.addImage({ path:FIG+"fig_families.png", x:0.4, y:1.28, w:9.2, h:3.55, sizing:{type:"contain",w:9.2,h:3.55} });
  bullets(s, ["fastText (same 1.2M words) is the BEST model at 100-300 labels, beats MLM at n=1000 (70.2 vs 65.7) - then plateaus at 71.2%",
              "VADER (0 task labels): 69.4% - worth ~840 TF-IDF labels",
              "Character n-grams: best typo robustness (78.7% at 30% typos)"], { x:0.45, y:4.9, w:9.1, h:0.7, fontSize:11.5 });
  note(s, "Slide 6 - Families", "4:00-5:00",
`Given the same unlabeled text, what's the best bet? It depends on how many labels you have. fastText embeddings, trained on the exact same words as our MLM run, are the best model overall at 100 and 300 labels, and still beat MLM pre-training at 1,000 labels - 70.2 versus 65.7 percent. But fastText plateaus around 71 percent while TF-IDF and the transformers keep climbing.
A sentiment lexicon with zero task labels reaches 69.4 percent - worth about 840 TF-IDF labels, a useful number when deciding whether to collect more labels at all.
And character n-grams give the best typo robustness of any classical model.`); }

// ================================================================ 7 Label-equivalence
{ const s = base("A currency for comparison: label-equivalence", "How many TF-IDF labels is each model worth? (interpolated on the TF-IDF curve)");
  const hdr = ["Model","n = 1,000","n = 6,920"].map(t=>({text:t,options:{bold:true,color:"FFFFFF",fill:{color:NAVY},fontFace:F,fontSize:12,align:"center"}}));
  const rows = [["VADER (0 labels)","~840","--"],["fastText (same words)","~960","~1,120"],["Tiny-BERT +MLM (ours)","~450","~3,000"],["BERT-Tiny (public)","~2,040",">6,920"],["DistilBERT (public)",">6,920",">6,920"]]
    .map((r,i)=>r.map((t,j)=>({text:t,options:{fontFace:F,fontSize:13,color:INK,align:j?"center":"left",fill:{color:i%2?"FFFFFF":LIGHT}}})));
  s.addTable([hdr,...rows], { x:0.45, y:1.35, w:9.1, colW:[4.1,2.5,2.5], rowH:0.4, border:{type:"solid",pt:0.5,color:"E5E7EB"} });
  stat(s, 0.45, 4.15, 4.4, "0.45x", "our pre-trained tiny BERT: 1,000 labels ~ 450 TF-IDF labels", ORANGE);
  stat(s, 5.15, 4.15, 4.4, "6.9x", "DistilBERT with 1,000 labels beats TF-IDF with all 6,920", BLUE);
  note(s, "Slide 7 - Label-equivalence", "5:00-5:30",
`To make these gaps concrete, I read every model's accuracy off the TF-IDF curve and asked: how many TF-IDF labels is this worth? Our pre-trained model at 1,000 labels is worth about 450 - less than half its own labels. DistilBERT at 1,000 labels beats TF-IDF with all 6,920 - almost a sevenfold saving. This already hints that scale, not the recipe, is what's missing.`); }

// ================================================================ 8 Decoupling
{ const s = base("RQ3: the probe improves, fine-tuning does not", "Checkpoints after 0-5,000 MLM steps; n = 1,000 labels; 3 seeds");
  s.addImage({ path:FIG+"fig_ckpt.png", x:0.45, y:1.3, w:5.1, h:3.55, sizing:{type:"contain",w:5.1,h:3.55} });
  stat(s, 5.85, 1.4, 3.7, "56.4 -> 59.4", "frozen linear probe: MLM does store sentiment info");
  stat(s, 5.85, 2.85, 3.7, "69.7 -> 65.7", "fine-tuned accuracy falls with pre-training", ORANGE);
  s.addText("A probe alone would have predicted a benefit.", { x:5.85, y:4.3, w:3.7, h:0.6, fontFace:F, fontSize:12.5, italic:true, color:MUTED, margin:0, isTextBox:true });
  note(s, "Slide 8 - Decoupling", "5:30-6:15",
`Why the negative result? I tested each pre-training checkpoint two ways: fine-tune it, or freeze it and train only a linear probe. They disagree. The probe rises from 56.4 to 59.4 percent - MLM does add sentiment-relevant information. But fine-tuned accuracy falls, from 69.7 to 65.7. Our pre-trained weights are a better representation but a worse starting point for fine-tuning. The next slide shows exactly where that plays out.`); }

// ================================================================ 9 Ambiguity
{ const s = base("RQ7: the deficit hides in plain sight", "Where does pre-training lose, and does confidence track difficulty?");
  s.addImage({ path:FIG+"fig_ambiguity.png", x:0.4, y:1.28, w:9.2, h:3.15, sizing:{type:"contain",w:9.2,h:3.15} });
  bullets(s, ["-4.7 pts on STRONG (clear) sentences; only -1.3 on WEAK (ambiguous) ones - the deficit is concentrated on the easy cases",
              "Best ambiguity-awareness AUROC across ALL models: 0.62 (chance = 0.5) - nobody's confidence tracks ambiguity well",
              "Transformers on 1,000 labels are badly overconfident: ECE 23-26% vs 3-6% for TF-IDF"], { x:0.45, y:4.5, w:9.1, h:1.0, fontSize:11.5 });
  note(s, "Slide 9 - Ambiguity", "6:15-7:15",
`Here's the twist. I split test sentences by polarity strength, using fine-grained SST-5 labels: strong, clearly positive or negative, versus weak, mild and closer to neutral. At 1,000 labels, pre-training costs 4.7 points on STRONG sentences but only 1.3 - not significant - on WEAK ones. The interaction is significant: pre-training's damage is concentrated on the sentences that were already easy.
And across every model I tried, confidence barely tracks ambiguity - the best AUROC for flagging a weak-polarity sentence from low confidence is 0.62, barely above chance. Worse, transformers trained on 1,000 labels are badly overconfident - calibration error of 23 to 26 percent, versus 3 to 6 for TF-IDF. You cannot trust a small transformer's confidence to tell you when it's unsure.`); }

// ================================================================ 10 Robustness
{ const s = base("RQ2: typos and domain shift", "Accuracy (%) at n = 6,920 labeled sentences");
  const labels = ["clean","typos 30%","OOD IMDb"];
  s.addChart(pres.charts.BAR, [
    {name:"TF-IDF + LogReg", labels, values:[81.3,77.5,72.5]},
    {name:"Tiny random init", labels, values:[78.5,73.6,68.1]},
    {name:"Tiny + MLM", labels, values:[77.5,70.6,63.7]},
    {name:"BERT-Tiny (public)", labels, values:[81.7,74.6,69.1]},
    {name:"DistilBERT (public)", labels, values:[90.4,83.9,72.8]}],
    Object.assign({}, chartBase, { x:0.35, y:1.3, w:6.2, h:3.7, barDir:"col", barGapWidthPct:50, chartColors:[GREY,ORANGE,TEAL,BLUE,PURPLE], valAxisMinVal:55, valAxisMaxVal:95, showValue:true, dataLabelFormatCode:"0.0", dataLabelFontSize:8, dataLabelColor:INK }));
  bullets(s, ["30% typos: transformers lose 4.9-7.1 pts vs 3.8 for TF-IDF",
              "OOD: pre-training did not help ours; no public model beat TF-IDF"], { x:6.7, y:1.45, w:2.9, h:3.5, fontSize:12.5 });
  note(s, "Slide 10 - Robustness", "7:15-8:00",
`Typos hurt every transformer more than TF-IDF - likely because WordPiece fragments misspelled words. And out-of-domain, on IMDb reviews, pre-training did not help our model, and no public checkpoint beat TF-IDF either.`); }

// ================================================================ 11 Twist
{ const s = base("The twist: pre-training is a scale phenomenon", "SST-2 accuracy (%), same fine-tuning code, 3 seeds");
  const labels = ["n = 1,000","n = 6,920"];
  s.addChart(pres.charts.BAR, [
    {name:"TF-IDF + LogReg", labels, values:[70.5,81.3]},
    {name:"Tiny + MLM (ours)", labels, values:[65.7,77.5]},
    {name:"BERT-Tiny (public, 4.4M)", labels, values:[75.0,81.7]},
    {name:"DistilBERT (public, 66.4M)", labels, values:[86.5,90.4]}],
    Object.assign({}, chartBase, { x:0.35, y:1.3, w:5.7, h:3.6, barDir:"col", barGapWidthPct:50, chartColors:[GREY,TEAL,BLUE,PURPLE], valAxisMinVal:55, valAxisMaxVal:95, showValue:true, dataLabelFormatCode:"0.0", dataLabelFontSize:8, dataLabelColor:INK }));
  bullets(s, ["BERT-Tiny has OUR depth and width: +9.3 pts at n=1k over our pre-trained model",
              "DistilBERT: +16/+9 pts over TF-IDF, ~26 min to fine-tune vs 37 s",
              "So: scale, not the recipe, is the missing ingredient"], { x:6.25, y:1.45, w:3.3, h:3.6, fontSize:12.5 });
  note(s, "Slide 11 - Twist", "8:00-8:45",
`If the architecture were the problem, public BERT-Tiny - same depth, same width - would fail too. It doesn't: 9.3 points above our pre-trained model at 1,000 labels. DistilBERT is further ahead still. So the negative result is about scale - 1.2 million words, 13 minutes - not about pre-training as a method.`); }

// ================================================================ 11b Cross-domain
{ const s = base("RQ8: does it travel? Cross-domain generalisation", "Financial PhraseBank (finance news, real agreement) + Twitter Airline (tweets, real confidence)");
  s.addImage({ path:FIG+"fig_cross_domain.png", x:0.35, y:1.25, w:9.3, h:3.3, sizing:{type:"contain",w:9.3,h:3.3} });
  bullets(s, ["Small-budget result REPLICATES: fastText beats +MLM by +3.2 pts [+2.3,+4.1] at n=100 (airline); +MLM significantly WORSE than random init (-1.9 [-2.5,-1.2])",
              "Full-data result does NOT replicate: TF-IDF's SST-2 edge vanishes - all 3 families tie once a domain has enough labels",
              "Cross-domain transfer costs 15-60 pts everywhere; TF-IDF travels best (62.7% movies->finance); pre-training does not close the gap"], { x:0.45, y:4.65, w:9.1, h:0.9, fontSize:11.5 });
  note(s, "Slide 11b - Cross-domain", "8:45-9:30",
`Does any of this generalise beyond movies? I repeated the two sharpest findings on two more domains with REAL ambiguity signals - finance news with actual annotator agreement tiers, and airline tweets with actual crowd-worker confidence scores.
The small-budget story replicates, even more sharply: on airline at 100 labels, fastText beats pre-training by 3.2 points, and pre-training is SIGNIFICANTLY WORSE than random init.
But the full-data story does not replicate. On SST-2, TF-IDF keeps a significant edge even at full data. In finance and airline, all three families - TF-IDF, random init, pre-trained - become statistically indistinguishable once there are enough labels. So TF-IDF's advantage was partly a movie-review-specific result, not a universal one.
And cross-domain transfer is uniformly poor - every model loses 15 to 60 points moving to a different domain, TF-IDF travels best, and pre-training does not help close that gap at all.`); }

// ================================================================ 12 Takeaways
{ const s = base("Takeaways", "Three findings, one field guide");
  stat(s, 0.45, 1.3, 2.9, "It depends", "fastText wins <1k labels; TF-IDF wins after; MLM never does", ORANGE);
  stat(s, 3.55, 1.3, 2.9, "Where it fails", "pre-training's damage sits on the EASY sentences", TEAL);
  stat(s, 6.65, 1.3, 2.9, "Don't trust it", "no model's confidence tracks ambiguity (AUROC <=0.62)", BLUE);
  bullets(s, ["Field guide: few hundred labels -> fastText; few thousand -> TF-IDF; public checkpoint available -> use it",
              "Scale, not the recipe, explains our negative result (BERT-Tiny, same architecture, does help)",
              "Next: more public checkpoints (code supports BERT-Mini/Small, MiniLM, TinyBERT, ELECTRA), tuned lr, second dataset"], { y:2.85, h:2.5, fontSize:13.5 });
  note(s, "Slide 12 - Takeaways", "8:45-10:00",
`Three findings. One: the best bet for unlabeled text depends on how many labels you have - fastText under a thousand, TF-IDF after, and our pre-training recipe never wins outright. Two: pre-training's damage concentrates on sentences that were already easy, and roughly disappears on ambiguous ones. Three: no model's confidence is a good ambiguity detector, and small transformers are dangerously overconfident.
Field guide: a few hundred labels, use fastText; a few thousand, TF-IDF; if you have a public checkpoint, use it. The whole pipeline is one command and fully reproducible. Thank you - happy to take questions.`); }

// ================================================================ BACKUP
{ const s = base("Backup: cost of each model", "Single CPU thread; seed-0 full-data models", true);
  const hdr = ["Model","Size (MB)","Latency (ms)","Sent./s","Acc. (%)"].map(t=>({text:t,options:{bold:true,color:"FFFFFF",fill:{color:NAVY},fontFace:F,fontSize:12,align:"center"}}));
  const rows = [["Tiny random init, FP32","4.71","1.03","4,914","78.3"],["Tiny +MLM, FP32","4.71","1.02","5,073","76.1"],
                ["Tiny +MLM, INT8 dynamic","3.54","1.71","3,196","75.9"],["TF-IDF + LogReg","2.72","0.31","48,388","81.3"],
                ["fastText + LogReg (est.)","~0.5","~0.5","~10,000","71.2"]]
    .map((r,i)=>r.map((t,j)=>({text:t,options:{fontFace:F,fontSize:12,color:INK,align:j?"center":"left",fill:{color:i%2?"FFFFFF":LIGHT}}})));
  s.addTable([hdr,...rows], { x:0.45, y:1.4, w:9.1, colW:[3.3,1.5,1.5,1.4,1.4], rowH:0.42, border:{type:"solid",pt:0.5,color:"E5E7EB"} });
  bullets(s, ["INT8: -25% file size, accuracy unchanged within noise (single model)", "fastText row is an order-of-magnitude estimate, not separately measured"], { y:3.75, h:1.4, fontSize:13 });
  note(s, "Backup B1 - Cost", "Q&A",
`Use if asked about efficiency. INT8 dynamic quantization cuts the tiny model's file by 25 percent with unchanged accuracy. TF-IDF is smaller, faster and more accurate than any transformer here. The fastText row is an estimate, not a separate measurement, and I say so if asked.`); }

{ const s = base("Backup: statistical evidence", "Paired bootstrap, 95% CI over test sentences, seeds pooled; * = CI excludes 0", true);
  const hdr = ["Comparison","n=1000","n=6920"].map(t=>({text:t,options:{bold:true,color:"FFFFFF",fill:{color:NAVY},fontFace:F,fontSize:12,align:"center"}}));
  const rows = [["+MLM - random init, STRONG sentences","-4.7* [-7.2,-2.2]","-2.6* [-4.8,-0.2]"],
                ["+MLM - random init, WEAK sentences","-1.3 [-3.4,+0.8]","-0.1 [-2.2,+1.9]"],
                ["+MLM - random init, TF-IDF LR (overall)","-2.5* [-4.2,-1.0]","-1.0 [-2.6,+0.5]"]]
    .map((r,i)=>r.map((t,j)=>({text:t,options:{fontFace:F,fontSize:12,color:INK,align:j?"center":"left",fill:{color:i%2?"FFFFFF":LIGHT}}})));
  s.addTable([hdr,...rows], { x:0.45, y:1.4, w:9.1, colW:[4.6,2.25,2.25], rowH:0.55, border:{type:"solid",pt:0.5,color:"E5E7EB"} });
  bullets(s, ["Interaction (strong - weak) at n=1000: -3.5 [-6.7,-0.3], CI excludes 0", "Method: resample 1,821 test sentences 2,000 times; correctness difference averaged over seeds"], { y:3.6, h:1.4, fontSize:13 });
  note(s, "Backup B2 - Statistics", "Q&A",
`The interaction test - is the strong-sentence gap bigger than the weak-sentence gap - has a 95 percent interval of -3.5 to -0.3, which excludes zero at n=1000. That's the statistical backbone of the ambiguity claim.`); }

{ const s = base("Backup: limitations and threats to validity", "What the results do and do not show", true);
  bullets(s, ["Ambiguity proxy is SST-5 polarity STRENGTH, not annotator disagreement; 'weak' covers 62.8% of sentences",
              "Only 2 public checkpoints (BERT-Tiny, DistilBERT), untuned learning rates, predictions not stored (no paired tests for them)",
              "Embedding families untuned (fixed dim/epochs/pooling); no CNN/BiLSTM baseline",
              "3 seeds; full 872-sentence dev set for early stopping even at n=100",
              "One task, one dataset, one language (SST-2, English)"], { fontSize:13.5 });
  note(s, "Backup B3 - Limitations", "Q&A",
`Be upfront: the ambiguity signal measures polarity strength, not real annotator disagreement, and covers most of the dataset as 'weak', so it's a coarse proxy. Public checkpoints used untuned learning rates and their predictions weren't stored, so no paired tests exist for them yet. Everything else is standard caveats: three seeds, one task, one language.`); }

{ const s = base("Backup: reproducibility", "Everything regenerates from logged runs", true);
  s.addShape(pres.ShapeType.rect,{x:0.45,y:1.4,w:9.1,h:2.5,fill:{color:"111827"},line:{color:"111827"}});
  s.addText([{text:"bash run_all.sh                      # clean run of the whole study",options:{breakLine:true}},
    {text:"python run_experiments.py --stage families  # VADER, SVM, char n-gram, word2vec/fastText",options:{breakLine:true}},
    {text:"python ambiguity.py                  # SST-5 strata, AUROC, ECE",options:{breakLine:true}},
    {text:"python hf_benchmark.py --models ladder --tune  # size ladder, tuned lr"}],
    {x:0.7,y:1.55,w:8.6,h:2.2,fontFace:"Consolas",fontSize:12.5,color:"E5E7EB",valign:"top",paraSpaceAfter:7,margin:0,isTextBox:true});
  bullets(s, ["Resumable, stored predictions, paired statistics throughout", "Deliverables: code, technical docs, main report, 2-page report, ACM paper, this deck"], { y:4.1, h:1.2, fontSize:13 });
  note(s, "Backup B4 - Reproducibility", "Q&A",
`One command runs everything. Every stage is resumable and stores its predictions, so statistics and the ambiguity analysis can always be regenerated or extended.`); }

pres.writeFile({ fileName: "/home/claude/tiny-nlp-research/slides/tiny_bert_pretraining_study.pptx" }).then(()=>console.log("written"));

// ================================================================ Notes + defense doc
const QA = [
["Is the negative result just a bug?", "No: MLM loss falls from 8.2 to 4.35 (works mechanically), the frozen probe improves with more pre-training (checkpoints really differ), and the same fine-tuning code gives strong results for public BERT-Tiny and DistilBERT. The claim is about scale, not a broken pipeline."],
["What is genuinely new here, since there's no new model?", "The contribution is the protocol and the analysis: one evaluation harness spanning a lexicon, static embeddings, TF-IDF, from-scratch and pre-trained transformers, and public checkpoints; the probe/fine-tune decoupling; and the ambiguity-stratified analysis showing WHERE pre-training's damage concentrates. It is an empirical study, not a new method, and I say so."],
["Why does fastText do well early and then plateau?", "Averaged static embeddings capture broad topical/sentiment signal with very few labels since the embedding space already groups similar words. But averaging loses word order and can't model negation or contrast well, so it can't keep improving as more labeled examples let sparse or contextual models fit finer decision boundaries. I did not verify this mechanism directly."],
["Is the ambiguity split (SST-5 strength) really 'ambiguity'?", "It's a graded-intensity proxy, not annotator disagreement, and I'm explicit about that. It's validated indirectly: sentences few models get right are disproportionately 'weak' polarity (79.6% vs 62.8% base rate), so it correlates with what's actually hard, but it is not the strongest possible ambiguity signal."],
["Why is the interaction significant at n=1000 but not n=6920?", "At n=1000 the strong-sentence gap is -4.7 and the weak-sentence gap is -1.3, a large and well-estimated difference. At n=6920 both models are more accurate and the gaps shrink (-2.6 and -0.1), so the difference is smaller and the confidence interval, while still negative on average, includes zero. More seeds would sharpen this."],
["Are three seeds enough?", "They give noisy standard deviations, which I flag as a limitation. The paired bootstrap covers test-set sampling variation independently. The main directional findings - TF-IDF ahead from 1,000 labels, pre-training's deficit on strong sentences - are consistent across multiple budgets and seeds, not resting on one comparison."],
["Why does the transformer do relatively well on 'contrast' sentences (but/although)?", "Contextual self-attention can in principle combine information across a sentence, so it may partially resolve a contrastive clause, whereas bag-of-words methods (TF-IDF, fastText) just average or sum features regardless of structure. The gap between the transformer and TF-IDF nearly disappears on contrast sentences (-0.1 vs -3.5 elsewhere). This is consistent with, not proof of, compositional processing; I did not run a targeted structural test."],
["Why are the transformers so overconfident at n=1000?", "With few examples, cross-entropy training can push output probabilities toward the extremes on the training set, and this miscalibration transfers to test time; it is a widely observed phenomenon in small-data fine-tuning, not unique to this study. TF-IDF's regularised linear scores stay closer to calibrated by construction."],
["Is the public-checkpoint comparison fair?", "Not perfectly: they used fixed learning rates (3e-4, 5e-5) while ours had a small tuned grid, which if anything understates their advantage. They also differ in vocabulary size from ours, which I can't separate from pre-training scale. I say this explicitly."],
["What would change your conclusion?", "If a larger in-domain pre-training run (more words, more steps) closed the gap to random init, that would show the negative result is purely about scale, matching what BERT-Tiny already suggests. If the ambiguity/strength split did not replicate with a real disagreement-based signal, I would weaken that claim specifically."],
["Why mean pooling and not the [CLS] token?", "MLM never trains the [CLS] token or a pooler head, so a randomly initialised pooler would add noise. Mean pooling works uniformly for random-init, pre-trained, and public models."],
["What's your practical recommendation?", "With a CPU, unlabeled in-domain text, and a few hundred labels: fastText. A few thousand: TF-IDF, word or character n-grams. If you have a public checkpoint: use it, and prefer DistilBERT-scale if you can afford the fine-tuning time. Never trust a small transformer's raw confidence to flag ambiguous inputs."],
["How reproducible is this?", "One command (bash run_all.sh) runs a full clean study; every stage is resumable, predictions are stored, and every figure/table regenerates from raw logs via make_assets.py. Different hardware shifts numbers by a fraction of a point; all numbers here are from one machine."],
["What's the single most interesting number?", "The interaction effect: -3.5 points [-6.7,-0.3] at n=1000, meaning pre-training's damage is concentrated on sentences that were easy to classify anyway, not spread evenly. That's the ambiguity twist in one number."],
["Could this be published somewhere?", "As a controlled empirical study with a negative result and an ambiguity-stratified analysis, yes - to a workshop or student-research track that accepts negative results and reproducibility studies, not as a novel-method paper at a top venue without a second dataset and more public-checkpoint coverage."]
];
let md = `# Speaker notes and defense guide\n\n**Talk:** 10 minutes (12 slides, ~50s each) + 5 minutes Q&A. Backup slides B1-B4 on demand only.\n\n## Timing plan\n\n| Slide | Time |\n|---|---|\n`;
NOTES.slice(0,12).forEach(n=>{ md += `| ${n[0]} | ${n[1]} |\n`; });
md += `\n## Speaker script\n\n`;
NOTES.forEach(n=>{ md += `### ${n[0]}  (${n[1]})\n\n${n[2]}\n\n`; });
md += `## Q&A defense bank (5 minutes)\n\n`;
QA.forEach((q,i)=>{ md += `**Q${i+1}. ${q[0]}**\n\n${q[1]}\n\n`; });
md += `## Key numbers\n\n- Model: 1,173,248 params; corpus 1.19M words; 5,000 steps, 13 min; final ppl 77, masked acc 29.1%.\n- SST-2 accuracy (n=1000/6920): TF-IDF LR 70.5/81.3; random init 68.2/78.5; +MLM 65.7/77.5; fastText 70.2/71.2; VADER 69.4 (any n); BERT-Tiny 75.0/81.7; DistilBERT 86.5/90.4.\n- MLM - random init: -0.3, -4.5*, -2.5*, -2.3*, -1.0 (n=100...6920).\n- Ambiguity (n=1000): strong -4.7* [-7.2,-2.2]; weak -1.3 [-3.4,+0.8]; interaction -3.5* [-6.7,-0.3].\n- Ambiguity AUROC (low conf -> weak polarity): 0.545-0.622, best = char TF-IDF at n=6920.\n- ECE at n=1000: TF-IDF 3.1-6.4%; transformers 22.9-25.6%.\n- Label-equivalence: VADER ~840; +MLM ours ~450 (n=1000); DistilBERT >6,920 (n=1000).\n`;
fs.writeFileSync("/home/claude/tiny-nlp-research/slides/speaker_notes_and_defense.md", md);
console.log("notes written", NOTES.length, QA.length);
