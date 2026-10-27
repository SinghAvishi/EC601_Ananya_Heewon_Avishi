function audit510k(pdfFile, xlsxFile)
%AUDIT510K  Scanner-disclosure audit of ONE FDA 510(k) Summary PDF -> Excel.
%
%   audit510k                      prompts (uigetfile / uiputfile) for the PDF and the output .xlsx
%   audit510k(pdfFile)             prompts only for the output location
%   audit510k(pdfFile, xlsxFile)   no prompts (useful for batch scripting)
%
%   PURPOSE
%   Pre-deployment scanner-compatibility research audit. The tool extracts product,
%   scanner, validation-dataset, performance and input-requirement information and
%   records, for EVERY field, a disclosure status, the supporting quotation and the
%   PDF page. It never infers missing values.
%
%   DISCLOSURE STATUSES
%     Disclosed / Partially disclosed / Not disclosed / Unclear / N/A
%   "Not disclosed" means: not found in the text of THIS public summary. It does NOT
%   establish that the manufacturer failed to collect or evaluate the information.
%
%   REQUIREMENTS
%   MATLAB R2016b or newer (R2019b+ recommended). No toolbox is strictly required, but
%   you need AT LEAST ONE working PDF text backend (tried in this order):
%     1. pdftotext (Poppler) on the system PATH, or set cfg.pdftotextPath  [best: keeps
%        table layout and gives true page breaks]
%     2. Python with the 'pypdf' package, configured through pyenv (R2019b+)
%     3. extractFileText (Text Analytics Toolbox) [page numbers only if the text
%        contains page breaks; otherwise pages are reported as unknown]
%   Excel output uses writetable (works without Excel installed).
%
%   HOW TO MODIFY EXTRACTION RULES
%   All keyword lists / regular expressions are in defaultConfig() (section headings,
%   scanner manufacturers, reconstruction terms, acquisition terms, metric names,
%   model stop-words). Field-specific logic lives in the rule*() functions.
%   Status logic is centralised in assignStatus().
%
%   IMPORTANT: scanner models, counts and metric values are located with heuristic
%   text patterns. Every scanner-model and performance row is flagged for manual
%   review. Treat the workbook as a traceable first pass, not as ground truth.

    cfg = defaultConfig();

    % ---------- 1. file selection ----------
    if nargin < 1 || isempty(pdfFile)
        [f, p] = uigetfile({'*.pdf', 'PDF files (*.pdf)'}, 'Select an FDA 510(k) Summary PDF');
        if isequal(f, 0), disp('No PDF selected. Aborting.'); return; end
        pdfFile = fullfile(p, f);
    end
    if ~exist(pdfFile, 'file')
        error('audit510k:noFile', 'PDF not found: %s', pdfFile);
    end
    if nargin < 2 || isempty(xlsxFile)
        [~, base] = fileparts(pdfFile);
        [f, p] = uiputfile({'*.xlsx', 'Excel workbook (*.xlsx)'}, 'Save Excel output as', [base '_510k_audit.xlsx']);
        if isequal(f, 0), disp('No output location chosen. Aborting.'); return; end
        xlsxFile = fullfile(p, f);
    end

    % ---------- 2. text extraction + quality gate ----------
    doc = extractPdfText(pdfFile, cfg);
    qc  = assessTextQuality(doc, cfg);
    if ~qc.ok
        error('audit510k:noUsableText', ...
            ['No usable text could be extracted, so NO workbook was written.\n' ...
             'Reason(s):\n  - %s\n' ...
             'If this PDF is a scanned image, OCR is required first (e.g., Adobe Acrobat "Recognize Text", ' ...
             'ocrmypdf, or Tesseract), then re-run on the OCR''d PDF.\n' ...
             'Backend log:\n  - %s'], ...
            strjoin(qc.messages, '\n  - '), strjoin([doc.warnings, {'(end)'}], '\n  - '));
    end

    corpus = buildCorpus(doc);
    corpus = segmentSections(corpus, cfg);
    M      = findScannerMentions(corpus, cfg);

    % ---------- 3. field extraction ----------
    res = {};
    [Rmod, modInfo] = ruleModality(corpus, cfg);
    res{end+1} = ruleDeviceName(corpus, cfg);
    res{end+1} = ruleManufacturer(corpus, cfg);
    res{end+1} = rule510k(corpus, cfg);
    res{end+1} = ruleClearanceDate(corpus, cfg);
    res{end+1} = Rmod;
    res{end+1} = ruleIntendedUse(corpus, cfg);

    [B1, B2] = ruleScannerSpecs(corpus, cfg, M);
    res{end+1} = B1;  res{end+1} = B2;
    res{end+1} = ruleParam(corpus, cfg, 'B3', 'Reconstruction algorithm(s)', cfg.paramPat.algorithm, 'algorithm');
    res{end+1} = ruleParam(corpus, cfg, 'B4', 'Reconstruction kernel(s)',   cfg.paramPat.kernel,    'kernel');
    res{end+1} = ruleParam(corpus, cfg, 'B5', 'Reconstruction strength/settings', cfg.paramPat.strength, 'strength');
    res{end+1} = ruleParam(corpus, cfg, 'B6', 'Slice thickness',            cfg.paramPat.slice,     'slice');
    res{end+1} = ruleAcquisition(corpus, cfg);

    res{end+1} = ruleCountField(corpus, cfg, 'C1', 'Total number of patients', cfg.pat.patients, 'year|old|\bage|month|week|day|\bmm\b|%');
    res{end+1} = ruleCountField(corpus, cfg, 'C2', 'Total studies / exams / scans / cases', cfg.pat.studies, 'year|old|\bage|month|week|day|\bmm\b|%');
    res{end+1} = ruleCountField(corpus, cfg, 'C2b', 'Total images / slices', cfg.pat.images, 'year|old|\bage|month|week|day|\bmm\b|%');
    res{end+1} = ruleCountField(corpus, cfg, 'C3', 'Number of clinical sites', cfg.pat.sites, 'year|old|month|week|day|\bmm\b|%');
    res{end+1} = ruleSiteSplits(corpus, cfg);
    [C5, C6, C7] = ruleValidationScanners(corpus, cfg, M);
    res{end+1} = C5;  res{end+1} = C6;  res{end+1} = C7;

    [D2, indep] = ruleIndependence(corpus, cfg);
    res{end+1} = D2;
    res{end+1} = ruleInternalExternal(corpus, cfg, indep);
    [D3, D4, D5, perfRows] = rulePerformance(corpus, cfg);
    res{end+1} = D3;  res{end+1} = D4;  res{end+1} = D5;

    res{end+1} = ruleInputType(corpus, cfg);
    res{end+1} = ruleAcqRequirements(corpus, cfg);
    res{end+1} = ruleSupportedScanners(corpus, cfg, M);
    res{end+1} = ruleLimitations(corpus, cfg);

    results = vertcat(res{:});
    [~, order] = sort({results.id});
    results = results(order);

    % ---------- 4. status post-processing + cross checks ----------
    results = applyNotApplicable(results, modInfo);
    [results, qcMsgs] = crossChecks(results, perfRows);
    for i = 1:numel(results)
        if strcmp(results(i).status, 'Not disclosed')
            results(i).note = strtrim([results(i).note ' ' cfg.notDisclosedNote]);
        end
    end

    % ---------- 5. export ----------
    runInfo = struct('pdf', pdfFile, 'backend', doc.backend, 'pageAware', doc.pageAware, ...
        'nPages', numel(doc.pages), 'warnings', {[doc.warnings, qc.messages, qcMsgs]}, ...
        'emptyPages', qc.emptyPages);
    exportWorkbook(xlsxFile, results, perfRows, M, cfg, runInfo);

    nRev = sum([results.review]);
    fprintf('\nDone. Wrote: %s\n', xlsxFile);
    fprintf('Backend: %s | pages: %d | fields: %d | flagged for manual review: %d\n', ...
        doc.backend, numel(doc.pages), numel(results), nRev);
    st = {results.status};
    for s = {'Disclosed','Partially disclosed','Not disclosed','Unclear','N/A'}
        fprintf('  %-20s %d\n', s{1}, sum(strcmp(st, s{1})));
    end
end

%% ======================================================================
%%  CONFIGURATION (edit rules here)
%% ======================================================================
function cfg = defaultConfig()
    cfg.pdftotextPath   = 'pdftotext';      % full path if not on PATH
    cfg.backends        = {'pdftotext', 'python', 'extractFileText'};
    cfg.minCharsPerPage = 80;               % pages with fewer non-space chars are "empty"
    cfg.maxEvidenceChars = 350;
    cfg.maxListItems    = 8;
    cfg.headerPages     = 4;                % label-style fields are sought here first
    cfg.maxIFUChars     = 1500;
    cfg.notDisclosedNote = ['Not found in the text of this public summary; this does not establish ' ...
                            'whether the manufacturer collected or evaluated it.'];

    % Context labels given to scanner mentions
    cfg.ctx.val   = 'Performance/validation section';
    cfg.ctx.near  = 'Dataset/validation language nearby (unconfirmed)';
    cfg.ctx.train = 'Training-data language nearby';
    cfg.ctx.desc  = 'Description/other (no validation link)';

    % Section headings: {name, regex}. Order = priority when one line matches several.
    cfg.sectionPatterns = {
        'IFU',         'indications?\s+for\s+use|intended\s+use|intended\s+purpose'
        'DEVDESC',     'device\s+description|description\s+of\s+(?:the\s+)?device|device\s+overview|product\s+description'
        'PREDICATE',   'predicate\s+devices?|substantial\s+equivalence|comparison\s+(?:to|with)\s+(?:the\s+)?predicate'
        'TECH',        'technological\s+characteristics|technology\s+comparison|comparison\s+of\s+technological'
        'PERFORMANCE', ['performance\s+(?:data|testing|evaluation|assessment|summary|characteristics)|' ...
                        '(?:non-?clinical|clinical|standalone|software|bench|algorithm)\s+(?:and\s+\w+\s+)?(?:performance|testing|validation|study|studies|data|evaluation)|' ...
                        'summary\s+of\s+(?:non-?clinical|clinical|performance)|validation(?:\s+(?:study|testing|data))?|study\s+design|test\s+dataset']
        'LIMITATIONS', 'limitations?|contraindications?|warnings?|precautions?'
        'CONCLUSION',  'conclusions?'
        'GENERAL',     'general\s+information|submitter(?:\s+information)?|administrative\s+information|510\s*\(k\)\s+summary|contact\s+information'
    };

    % Scanner manufacturers: {canonical name, regex alternatives (longest first)}
    cfg.mfrPatterns = {
        'GE Healthcare',  'GE\s+Healthcare|GE\s+Medical\s+Systems|General\s+Electric|GE'
        'Siemens',        'Siemens(?:\s+Healthineers|\s+Healthcare|\s+Medical\s+Solutions)?'
        'Philips',        'Philips(?:\s+Healthcare|\s+Medical\s+Systems)?'
        'Canon',          'Canon(?:\s+Medical\s+Systems)?'
        'Toshiba',        'Toshiba(?:\s+Medical\s+Systems)?'
        'Hologic',        'Hologic'
        'FUJIFILM',       'Fuji\s*film(?:\s+Healthcare)?'
        'United Imaging', 'United\s+Imaging(?:\s+Healthcare)?'
        'Carestream',     'Carestream'
        'Samsung',        'Samsung'
        'Neusoft',        'Neusoft'
        'Hitachi',        'Hitachi'
        'Varian',         'Varian'
        'Konica Minolta', 'Konica\s+Minolta'
        'Agfa',           'Agfa'
        'Shimadzu',       'Shimadzu'
        'Planmed',        'Planmed'
    };

    % Words that end a scanner-model phrase (lower case)
    cfg.modelStop = {'and','or','with','using','scanner','scanners','system','systems','were','was','from', ...
        'for','the','in','of','on','at','as','to','is','are','that','which','by','manufactured','manufacturer', ...
        'manufacturers','models','model','vendors','vendor','platform','platforms','device','devices','equipment', ...
        'imaging','data','images','image','cases','patients','studies','study','mri','ct','mr','inc','gmbh','corp', ...
        'ltd','llc','co','ag','healthcare','medical','healthineers','be','also','include','includes','including', ...
        'such','other','others','various','different','multiple','several','each','all','scans','scan','acquired', ...
        'acquisitions','sites','site','a','an','this','these','those','their','its','it','vendor-specific'};

    % Reconstruction / acquisition patterns (applied with ignorecase)
    cfg.paramPat.algorithm = ['\b(?:filtered\s+back[\s-]?projection|FBP|(?:adaptive\s+statistical\s+)?iterative\s+reconstruction|' ...
        'ASiR-?V|ASiR|Veo|MBIR|SAFIRE|ADMIRE|iDose\s*\d*|IMR|AIDR\s*3D(?:\s+Enhanced)?|AIDR|AiCE|DLIR|TrueFidelity|' ...
        'Precise\s+Image|deep[\s-]learning\s+(?:image\s+)?reconstruction)\b'];
    cfg.paramPat.algorithmSpecific = 'ASiR|Veo|MBIR|SAFIRE|ADMIRE|iDose|IMR|AIDR|AiCE|DLIR|TrueFidelity|Precise|FBP|filtered';
    cfg.paramPat.kernel    = '\b(?:convolution\s+|reconstruction\s+)?kernels?\b[^.;\r\n]{0,120}';
    cfg.paramPat.kernelName = '\b[A-Za-z]{1,3}\d{2}[a-z]?\b|\b(?:standard|soft|sharp|smooth|bone|lung|detail|body|chest|brain|medium)\b';
    cfg.paramPat.strength  = ['\b(?:ASiR-?V|ASIR|SAFIRE|ADMIRE|iDose|AIDR\s*3D|DLIR|AiCE|IMR|Veo)\b[^.;\r\n]{0,40}?' ...
        '(?:\d{1,3}\s*%|(?:strength|level)\s*[:=]?\s*\w+|\b(?:low|mild|medium|moderate|high|strong|standard)\b)|' ...
        '\b(?:strength|blending|blend)\s+(?:level\s+)?(?:of\s+)?\d{1,3}\s*%?'];
    cfg.paramPat.slice     = ['slice\s+thickness(?:es)?[^.;\r\n]{0,60}?\d+(?:\.\d+)?\s*(?:-|to)?\s*(?:\d+(?:\.\d+)?)?\s*mm|' ...
        '\d+(?:\.\d+)?\s*(?:-|to)?\s*(?:\d+(?:\.\d+)?)?\s*mm\s+(?:thick|slice|axial|section)|' ...
        '(?:thickness|slices?)\s*(?:of|<=|>=|<|>|=)\s*\d+(?:\.\d+)?\s*mm'];
    cfg.acqPatterns = {
        'Tube voltage',      '\b\d{2,3}(?:\s*-\s*\d{2,3})?\s*kVp?\b|\b(?:kVp|tube\s+voltage)\b[^.;\r\n]{0,40}'
        'Tube current/mAs',  '\b\d+(?:\.\d+)?\s*(?:-\s*\d+(?:\.\d+)?\s*)?mAs?\b|\btube\s+current\b[^.;\r\n]{0,40}|\bmAs\b[^.;\r\n]{0,30}'
        'Pitch',             '\bpitch\b[^.;\r\n]{0,30}'
        'Collimation',       '\bcollimation\b[^.;\r\n]{0,40}'
        'Field of view',     '\b(?:field[\s-]of[\s-]view|FOV)\b[^.;\r\n]{0,40}'
        'Contrast',          '\b(?:non-?contrast|(?:with|without)\s+(?:IV\s+|intravenous\s+)?contrast|contrast[\s-]enhanced|unenhanced)\b'
        'MR field strength', '\b(?:1\.5|3(?:\.0)?|7)\s*(?:T|Tesla)\b'
        'MR sequence',       '\b(?:T1|T2|FLAIR|DWI|SWI|MPRAGE|STIR)[\w\-]*(?:[\s-]weighted)?\b'
        'TE/TR/TI',          '\b(?:TE|TR|TI)\s*[=:]\s*\d+[^.;\r\n]{0,15}'
    };

    % Modalities: {label, regex}
    cfg.modalities = {
        'CT',          '\bCT\b|computed\s+tomography'
        'MRI',         '\bMRI\b|magnetic\s+resonance|\bMR\s+(?:image|images|scanner|system)'
        'X-ray',       'x-?\s?ray|radiograph|\bCXR\b'
        'Mammography', 'mammograph|tomosynthesis|\bDBT\b'
        'Ultrasound',  'ultrasound|sonograph'
        'PET/SPECT',   '\bPET\b|\bSPECT\b|positron|nuclear\s+medicine'
        'Fluoroscopy', 'fluorosc'
    };

    % Counting patterns: {regex, default unit}. Group 1 = number (digits or word), group 2 (optional) = unit.
    num  = '(\d{1,3}(?:,\d{3})+|\d+)';
    nw   = '(\d+|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen)';
    siteU = '(sites?|institutions?|hospitals?|centers?|centres?|facilities|health\s+systems?|practices)';
    cfg.pat.patients = {
        ['\b' num '\s+(?:[A-Za-z\-]+\s+){0,2}?(patients?|subjects?|participants?|individuals?)\b'], 'patients'
        '\b(?:patients?|subjects?|participants?)\s*\(\s*n\s*=\s*(\d[\d,]*)\s*\)',                    'patients'
        '\bn\s*=\s*(\d[\d,]*)\s*(?:patients?|subjects?|participants?)\b',                            'patients'};
    cfg.pat.studies = {
        ['\b' num '\s+(?:[A-Za-z\-]+\s+){0,2}?(studies|exams?|examinations?|scans?|cases|series|volumes?|radiographs?|mammograms?)\b'], 'studies'
        '\b(?:studies|exams|scans|cases)\s*\(\s*n\s*=\s*(\d[\d,]*)\s*\)',                                                         'studies'};
    cfg.pat.images = {
        ['\b' num '\s+(?:[A-Za-z\-]+\s+){0,2}?(images?|slices?)\b'], 'images'
        '\b(?:images|slices)\s*\(\s*n\s*=\s*(\d[\d,]*)\s*\)',        'images'};
    cfg.pat.sites = {
        ['\b' nw '\s+(?:[A-Za-z\-]+\s+){0,3}?(?:clinical\s+)?' siteU '\b'], 'sites'
        '\b(?:sites?|institutions?|centers?|hospitals?)\s*\(\s*n\s*=\s*(\d+)\s*\)', 'sites'};
    cfg.nw = nw;  cfg.siteU = siteU;

    % Performance metrics: {label, regex}
    cfg.metrics = {
        'Sensitivity',  '\bsensitivity\b|\brecall\b|true\s+positive\s+rate'
        'Specificity',  '\bspecificity\b|true\s+negative\s+rate'
        'AUC',          '\bAUC\b|\bAUROC\b|area\s+under\s+the\s+(?:receiver[^.\r\n]{0,40})?curve'
        'Accuracy',     '\baccuracy\b'
        'PPV',          '\bPPV\b|positive\s+predictive\s+value|\bprecision\b'
        'NPV',          '\bNPV\b|negative\s+predictive\s+value'
        'F1',           '\bF1(?:[\s-]score)?\b'
        'Dice',         '\bDice\b'
        'Kappa/ICC',    '\bkappa\b|\bICC\b|intraclass'
        'FP rate',      'false[\s-]positive\s+(?:rate|per)|false\s+positives?\s+per'
        'Error (MAE/RMSE)', '\bMAE\b|\bRMSE\b|mean\s+absolute\s+error'
        'Time-to-notification', 'time[\s-]to[\s-](?:notification|alert|triage)'
    };
end

%% ======================================================================
%%  PDF TEXT EXTRACTION
%% ======================================================================
function doc = extractPdfText(pdfFile, cfg)
% Tries each backend in cfg.backends; returns normalized page texts.
    doc = struct('pages', {{}}, 'backend', 'none', 'pageAware', false, 'warnings', {{}});
    for b = 1:numel(cfg.backends)
        name = cfg.backends{b};
        try
            switch name
                case 'pdftotext',       out = tryPdftotext(pdfFile, cfg);
                case 'python',          out = tryPython(pdfFile);
                case 'extractFileText', out = tryExtractFileText(pdfFile);
                otherwise,              continue;
            end
        catch ME
            out = struct('pages', {{}}, 'ok', false, 'msg', ME.message, 'pageAware', false);
        end
        hasText = out.ok && any(cellfun(@(t) ~isempty(strtrim(t)), out.pages));
        if hasText
            doc.pages     = cellfun(@normalizeText, out.pages, 'UniformOutput', false);
            doc.backend   = name;
            doc.pageAware = out.pageAware;
            if ~out.pageAware
                doc.warnings{end+1} = ['Backend ''' name ''' provided no page breaks: ALL page numbers are reported as unknown.'];
            end
            return;
        else
            msg = out.msg; if isempty(msg), msg = 'no text returned'; end
            doc.warnings{end+1} = sprintf('Backend %s failed: %s', name, msg);
        end
    end
end

function out = tryPdftotext(pdfFile, cfg)
    out = struct('pages', {{}}, 'ok', false, 'msg', '', 'pageAware', true);
    tmp = [tempname '.txt'];
    cmd = sprintf('"%s" -layout -enc UTF-8 "%s" "%s"', cfg.pdftotextPath, pdfFile, tmp);
    if ispc, cmd = ['"' cmd '"']; end
    [status, msg] = system(cmd);
    if status ~= 0 || ~exist(tmp, 'file')
        out.msg = ['pdftotext unavailable or failed: ' strtrim(msg)];
        return;
    end
    raw = fileread(tmp);
    delete(tmp);
    pg = strsplit(raw, char(12), 'CollapseDelimiters', false);
    if ~isempty(raw) && raw(end) == char(12), pg(end) = []; end   % pdftotext ends every page with FF
    out.pages = pg;  out.ok = ~isempty(pg);
end

function out = tryPython(pdfFile)
    out = struct('pages', {{}}, 'ok', false, 'msg', '', 'pageAware', true);
    try
        pe = pyenv;
        if isempty(char(pe.Version))
            out.msg = 'No Python environment configured (pyenv).'; return;
        end
        mod = py.importlib.import_module('pypdf');
        reader = mod.PdfReader(pdfFile);
        pgs = py.list(reader.pages);
        n = double(py.len(pgs));
        pages = cell(1, n);
        for k = 1:n
            pg = pgs{k};
            try
                t = pg.extract_text(pyargs('extraction_mode', 'layout'));
            catch
                t = pg.extract_text();
            end
            try
                pages{k} = char(t);
            catch
                pages{k} = '';
            end
        end
        out.pages = pages;  out.ok = n > 0;
    catch ME
        out.msg = ME.message;
    end
end

function out = tryExtractFileText(pdfFile)
    out = struct('pages', {{}}, 'ok', false, 'msg', '', 'pageAware', false);
    if exist('extractFileText', 'file') == 0
        out.msg = 'extractFileText not available (Text Analytics Toolbox).'; return;
    end
    txt = char(extractFileText(pdfFile));
    pg = strsplit(txt, char(12), 'CollapseDelimiters', false);
    out.pageAware = numel(pg) > 1;
    out.pages = pg;  out.ok = ~isempty(strtrim(txt));
end

function t = normalizeText(t)
% Light normalization only (keeps wording intact for quotations).
    t = char(t);
    t = strrep(t, char(13), '');
    t = strrep(t, char(160), ' ');
    t = strrep(t, char(64257), 'fi');  t = strrep(t, char(64258), 'fl');
    t = regexprep(t, ['[' char(8208) char(8209) char(8210) char(8211) char(8212) char(8722) ']'], '-');
    t = regexprep(t, ['[' char(8216) char(8217) ']'], '''');
    t = regexprep(t, ['[' char(8220) char(8221) ']'], '"');
    t = strrep(t, char(8804), '<=');  t = strrep(t, char(8805), '>=');
    t = strrep(t, char(8226), '*');
    t = strrep(t, char(0), '');
    t = strrep(t, char(9), '    ');
end

function qc = assessTextQuality(doc, cfg)
% Gate: refuse to continue on scanned / garbled PDFs rather than export an empty workbook.
    qc = struct('ok', true, 'messages', {{}}, 'emptyPages', []);
    n = numel(doc.pages);
    if n == 0
        qc.ok = false; qc.messages{end+1} = 'No text from any extraction backend.'; return;
    end
    nChars = cellfun(@(t) numel(regexprep(t, '\s', '')), doc.pages);
    qc.emptyPages = find(nChars < cfg.minCharsPerPage);
    allText = strjoin(doc.pages, ' ');
    alphaRatio = sum(isletter(allText)) / max(sum(~isspace(allText)), 1);
    if sum(nChars) < 500
        qc.ok = false; qc.messages{end+1} = sprintf('Only %d non-space characters in %d page(s): likely a scanned/image-only PDF.', sum(nChars), n);
    end
    if numel(qc.emptyPages) >= 0.5 * n
        qc.ok = false; qc.messages{end+1} = sprintf('%d of %d pages contain almost no text (likely scanned).', numel(qc.emptyPages), n);
    end
    if alphaRatio < 0.5
        qc.ok = false; qc.messages{end+1} = sprintf('Extracted text looks garbled (letters = %.0f%% of characters).', 100*alphaRatio);
    end
    if qc.ok
        if ~isempty(qc.emptyPages)
            qc.messages{end+1} = ['WARNING: pages with little/no extractable text (possible image-only tables/figures; OCR may be required for these): ' mat2str(qc.emptyPages)];
        end
        if isempty(regexpi(allText, '510\s*\(k\)|substantial\w*\s+equivalen|Food\s+and\s+Drug\s+Administration|indications?\s+for\s+use', 'once'))
            qc.messages{end+1} = 'WARNING: none of the usual 510(k) summary phrases were found; verify that this is a 510(k) Summary.';
        end
    end
end

%% ======================================================================
%%  CORPUS, SECTIONS, SEARCH HELPERS
%% ======================================================================
function corpus = buildCorpus(doc)
% One long string; pageStart(i) = index of first character of page i.
    n = numel(doc.pages);
    starts = zeros(1, n);  parts = cell(1, n);  pos = 1;  nl = char(10);
    for i = 1:n
        starts(i) = pos;
        parts{i} = [doc.pages{i} nl nl];
        pos = pos + numel(parts{i});
    end
    corpus.text      = [parts{:}];
    corpus.pageStart = starts;
    corpus.nPages    = n;
    corpus.pageAware = doc.pageAware;
    corpus.sections  = struct('name', {}, 'start', {}, 'stop', {}, 'heading', {});
end

function corpus = segmentSections(corpus, cfg)
% Heading heuristics: a short line consisting of a known heading (optionally numbered,
% optionally followed by a colon + text) starts a new section. Edit cfg.sectionPatterns.
    G = corpus.text;
    marks = struct('name', {}, 'start', {}, 'heading', {});
    for k = 1:size(cfg.sectionPatterns, 1)
        pat = ['^[ \t]*(?:(?:\d+(?:\.\d+)*|[IVXivx]+|[A-Za-z])[\.\)][ \t]+)?(?:' cfg.sectionPatterns{k,2} ')' ...
               '(?:[ \t]*[/&,\-]?[ \t]*[A-Za-z]+){0,4}(?:[ \t]*\([^)\r\n]{0,60}\))?[ \t]*(?::|$)'];
        [s, m] = regexp(G, pat, 'start', 'match', 'lineanchors', 'ignorecase');
        for i = 1:numel(s)
            h = strtrim(m{i});
            if numel(h) > 90, continue; end
            if ~isempty(marks) && any([marks.start] == s(i)), continue; end
            marks(end+1) = struct('name', cfg.sectionPatterns{k,1}, 'start', s(i), 'heading', h); %#ok<AGROW>
        end
    end
    if ~isempty(marks)
        [~, ix] = sort([marks.start]);  marks = marks(ix);
        for i = 1:numel(marks)
            if i < numel(marks), stop = marks(i+1).start - 1; else, stop = numel(G); end
            corpus.sections(end+1) = struct('name', marks(i).name, 'start', marks(i).start, ...
                                            'stop', stop, 'heading', marks(i).heading); %#ok<AGROW>
        end
    end
end

function name = sectionAt(corpus, pos)
    name = 'PREAMBLE';
    for i = numel(corpus.sections):-1:1
        if corpus.sections(i).start <= pos, name = corpus.sections(i).name; return; end
    end
end

function pg = pageOf(corpus, pos)
    if ~corpus.pageAware, pg = NaN; else, pg = find(corpus.pageStart <= pos, 1, 'last'); end
end

function s = pageStr(pg)
    if isnan(pg), s = '?'; else, s = sprintf('%d', pg); end
end

function s = pagesToStr(pgs)
    pgs = unique(pgs(~isnan(pgs)));
    if isempty(pgs), s = 'unknown'; else, s = strjoin(arrayfun(@(x) sprintf('%d', x), pgs, 'UniformOutput', false), ', '); end
end

function hits = findAll(corpus, pattern, caseSensitive)
% All regex matches with position, tokens, page and section.
    if nargin < 3, caseSensitive = false; end
    opts = {'tokens', 'match', 'start', 'end'};
    if ~caseSensitive, opts{end+1} = 'ignorecase'; end
    [tok, mat, s, e] = regexp(corpus.text, pattern, opts{:});
    hits = struct('s', {}, 'e', {}, 'match', {}, 'tok', {}, 'page', {}, 'section', {});
    for i = 1:numel(s)
        hits(end+1) = struct('s', s(i), 'e', e(i), 'match', mat{i}, 'tok', {tok{i}}, ...
                             'page', pageOf(corpus, s(i)), 'section', sectionAt(corpus, s(i))); %#ok<AGROW>
    end
end

function q = quoteAround(corpus, s, e, pre, post, cfg)
    a = max(1, s - pre);  b = min(numel(corpus.text), e + post);
    q = strtrim(regexprep(corpus.text(a:b), '\s+', ' '));
    if numel(q) > cfg.maxEvidenceChars, q = [q(1:cfg.maxEvidenceChars) ' ...']; end
end

function [h, late] = preferEarly(hits, corpus, cfg)
% Label-style fields (name, manufacturer) normally sit in the first pages.
    late = false;
    if isempty(hits), h = hits; return; end
    if corpus.pageAware
        early = [hits.page] <= cfg.headerPages;
    else
        early = [hits.s] <= 6000 * cfg.headerPages;
    end
    if any(early), h = hits(early); else, h = hits; late = true; end
end

function [h, fallback] = preferSection(hits, names)
    fallback = false;
    if isempty(hits), h = hits; return; end
    in = ismember({hits.section}, names);
    if any(in), h = hits(in); else, h = hits; fallback = true; end
end

function hits = dropNegated(hits, corpus)
% Drop hits preceded (same sentence, 40 chars) by no / not / without / never.
    keep = true(1, numel(hits));
    for i = 1:numel(hits)
        pre = corpus.text(max(1, hits(i).s - 40):hits(i).s - 1);
        if ~isempty(regexpi(pre, '(?:\bno\b|\bnot\b|\bwithout\b|\bnever\b)[^.]*$', 'once')), keep(i) = false; end
    end
    hits = hits(keep);
end

function [hits, vals] = dedupHits(hits, vals)
    if isempty(hits), return; end
    [~, ia] = unique(lower(vals), 'stable');
    hits = hits(ia);  vals = vals(ia);
end

function u = uniqueStable(c)
    if isempty(c), u = {}; return; end
    [~, ia] = unique(c, 'stable');
    u = reshape(c(ia), 1, []);
end

function s = joinList(c, d)
    if isempty(c), s = ''; else, c = reshape(c, 1, []); s = strjoin(c, d); end
end

function v = cleanLabelValue(v)
    v = regexprep(v, '\s{3,}.*$', '');          % drop adjacent table column
    v = strtrim(v);
    v = regexprep(v, '[\s\.,;:]+$', '');
end

function d = makeDetails(label, hits, vals, corpus, cfg)
    d = struct('label', {}, 'value', {}, 'page', {}, 'section', {}, 'quote', {});
    for i = 1:numel(hits)
        d(end+1) = struct('label', label, 'value', vals{i}, 'page', pageStr(hits(i).page), ...
            'section', hits(i).section, 'quote', quoteAround(corpus, hits(i).s, hits(i).e, 60, 100, cfg)); %#ok<AGROW>
    end
end

%% ======================================================================
%%  RESULT RECORDS + DISCLOSURE STATUS
%% ======================================================================
function R = newResult(id, group, field)
    R = struct('id', id, 'group', group, 'field', field, 'value', '', 'status', 'Not disclosed', ...
               'evidence', '', 'page', '', 'section', '', 'review', false, 'note', '', 'details', []);
end

function st = assignStatus(found, complete, ambiguous, notApplicable)
% Central status logic.
%   notApplicable -> N/A | ambiguous -> Unclear | ~found -> Not disclosed
%   found & complete -> Disclosed | found & ~complete -> Partially disclosed
    if notApplicable,   st = 'N/A';
    elseif ambiguous,   st = 'Unclear';
    elseif ~found,      st = 'Not disclosed';
    elseif complete,    st = 'Disclosed';
    else,               st = 'Partially disclosed';
    end
end

function R = fillFromHits(R, hits, vals, corpus, cfg)
% Value = distinct matched strings; evidence = up to 3 quotations with page numbers.
    if isempty(hits), return; end
    [uv, ia] = unique(vals, 'stable');
    nv = min(numel(uv), cfg.maxListItems);
    R.value = joinList(uv(1:nv), '; ');
    if numel(uv) > nv, R.value = [R.value sprintf('; ... (+%d more)', numel(uv) - nv)]; end
    ne = min(numel(ia), 3);
    ev = cell(1, ne);  pg = zeros(1, ne);
    for k = 1:ne
        h = hits(ia(k));
        ev{k} = sprintf('[p.%s] "%s"', pageStr(h.page), quoteAround(corpus, h.s, h.e, 50, 120, cfg));
        pg(k) = h.page;
    end
    R.evidence = joinList(ev, ' || ');
    R.page     = pagesToStr(pg);
    R.section  = hits(ia(1)).section;
end

%% ======================================================================
%%  A. PRODUCT IDENTIFICATION
%% ======================================================================
function R = ruleDeviceName(corpus, cfg)
    R = newResult('A1', 'A. Product identification', 'Device/product name');
    pat = ['(?:Trade\s*/\s*Device\s+Name|Device\s+Trade\s+Name|Trade\s+or\s+Proprietary\s+Name|Proprietary\s+Name|' ...
           'Name\s+of\s+Device|Trade\s+Name|Device\s+Name)\s*(?::[ \t]*|[ \t]{2,})([^\r\n]{2,150})'];
    hits = findAll(corpus, pat);
    if ~isempty(hits)
        hits = preferEarly(hits, corpus, cfg);
        vals = cellfun(@(t) cleanLabelValue(t{1}), {hits.tok}, 'UniformOutput', false);
        ok = cellfun(@(v) numel(v) >= 2 && any(isletter(v)), vals);
        hits = hits(ok);  vals = vals(ok);
    end
    if isempty(hits)
        R.review = true;  R.note = 'No device-name label found; check page 1 manually.';  return;
    end
    R = fillFromHits(R, hits(1), vals(1), corpus, cfg);
    nxt = regexp(corpus.text(hits(1).e+1:min(end, hits(1).e+200)), '^\n[ \t]{6,}([^\n:]{2,100})(?:\n|$)', 'tokens', 'once');
    if ~isempty(nxt)
        R.value = [R.value ' ' strtrim(nxt{1})];
        R.review = true;  R.note = 'Possible wrapped (multi-line) name appended; verify.';
    end
    uv = uniqueStable(lower(vals));
    if numel(uv) > 1
        R.review = true;
        R.note = strtrim([R.note ' CONFLICT/VARIANTS: ' joinList(uniqueStable(vals), ' | ')]);
    end
    if numel(R.value) > 120
        R.review = true;  R.note = strtrim([R.note ' Unusually long name; verify.']);
    end
    R.status = assignStatus(true, true, false, false);
end

function R = ruleManufacturer(corpus, cfg)
    R = newResult('A2', 'A. Product identification', 'Manufacturer');
    pat = ['(?:Submitter(?:''s)?(?:\s+Name)?|Applicant(?:\s+Name)?|510\s*\(k\)\s+(?:Owner|Sponsor|Submitter)|' ...
           'Sponsor|Company\s+Name|Manufacturer(?:\s+Name)?)\s*:\s*([^\r\n]{2,120})'];
    hits = findAll(corpus, pat);
    late = false;
    if ~isempty(hits)
        [hits, late] = preferEarly(hits, corpus, cfg);
        vals = cellfun(@(t) cleanLabelValue(t{1}), {hits.tok}, 'UniformOutput', false);
        ok = cellfun(@(v) numel(v) >= 2 && any(isletter(v)), vals);
        hits = hits(ok);  vals = vals(ok);
    end
    if isempty(hits)
        R.review = true;  R.note = 'No submitter/manufacturer label found; check page 1 and the FDA letter manually.';  return;
    end
    R = fillFromHits(R, hits(1), vals(1), corpus, cfg);
    R.status = assignStatus(true, true, false, false);
    R.review = true;
    R.note = 'Value is the text after the label (may include address); the label may denote the submitter rather than the manufacturer.';
    if late, R.note = [R.note ' Label found only after the first pages.']; end
    if numel(uniqueStable(lower(vals))) > 1
        R.note = [R.note ' VARIANTS: ' joinList(uniqueStable(vals), ' | ')];
    end
end

function R = rule510k(corpus, cfg)
    R = newResult('A3', 'A. Product identification', '510(k) number');
    numPat = '(K\d{6}|DEN\d{6})';
    lab = findAll(corpus, ['(?:510\s*\(k\)\s*(?:Number|No\.?|#)|Submission\s+Number|Premarket\s+Notification(?:\s+Number)?|\bRe)\s*:\s*' numPat]);
    labB = findAll(corpus, ['510\s*\(k\)\s+Summary[^\r\n]{0,60}?' numPat]);
    labHits = [lab, labB];
    anyNum = findAll(corpus, ['\b' numPat '\b'], true);
    allNums = uniqueStable(cellfun(@(t) t{1}, {anyNum.tok}, 'UniformOutput', false));
    if ~isempty(labHits)
        vals = upper(cellfun(@(t) t{end}, {labHits.tok}, 'UniformOutput', false));
        uvals = uniqueStable(vals);
        R = fillFromHits(R, labHits(1), vals(1), corpus, cfg);
        amb = numel(uvals) > 1;
        if amb
            R.value = joinList(uvals, '; ');
            R.review = true;  R.note = 'CONFLICT: multiple labelled 510(k)-format numbers.';
        end
        others = setdiff(allNums, uvals);
        if ~isempty(others)
            R.note = strtrim([R.note ' Other 510(k)-format numbers in document (likely predicate/reference devices; verify): ' joinList(others, ', ')]);
        end
        [ok, msg] = plausible510k(vals{1});
        if ~ok, R.review = true;  R.note = strtrim([R.note ' ' msg]); end
        R.status = assignStatus(true, true, amb, false);
    elseif ~isempty(allNums)
        R.value = '';
        R.evidence = sprintf('[p.%s] "%s"', pageStr(anyNum(1).page), quoteAround(corpus, anyNum(1).s, anyNum(1).e, 60, 80, cfg));
        R.page = pagesToStr(anyNum(1).page);  R.section = anyNum(1).section;
        R.review = true;
        R.note = ['Number(s) in 510(k) format appear but none is labelled as this submission (could be predicate devices): ' joinList(allNums, ', ')];
        R.status = assignStatus(true, false, true, false);
    else
        R.review = true;  R.note = 'No K######/DEN###### pattern found.';
    end
end

function [ok, msg] = plausible510k(num)
    ok = true;  msg = '';
    if strncmp(num, 'DEN', 3), yy = str2double(num(4:5)); else, yy = str2double(num(2:3)); end
    c = clock;  cur = mod(c(1), 100);
    if ~(yy >= 76 || yy <= cur)
        ok = false;  msg = sprintf('Implausible 510(k) number year prefix (%02d).', yy);
    end
end

function R = ruleClearanceDate(corpus, cfg)
    R = newResult('A4', 'A. Product identification', 'Clearance date');
    months = 'January|February|March|April|May|June|July|August|September|October|November|December';
    dateRx = ['(?:(?:' months ')\s+\d{1,2},?\s+\d{4}|\d{1,2}/\d{1,2}/\d{2,4}|\d{4}[./-]\d{2}[./-]\d{2})'];
    lab = findAll(corpus, ['(?:Decision\s+Date|Clearance\s+Date|Date\s+of\s+(?:Clearance|Decision)|Date\s+Cleared)\s*:?\s*(' dateRx ')']);
    if ~isempty(lab)
        R = fillFromHits(R, lab(1), {lab(1).tok{1}}, corpus, cfg);
        R.status = assignStatus(true, true, false, false);
        return;
    end
    % Date on the FDA clearance letter (precedes "Dear ..."; letter contains "substantially equivalent")
    dear = findAll(corpus, 'Dear\s+\w+');
    for i = 1:numel(dear)
        after = corpus.text(dear(i).e:min(end, dear(i).e + 3000));
        if isempty(regexpi(after, 'substantially\s+equivalent', 'once')), continue; end
        a = max(1, dear(i).s - 1500);
        w = corpus.text(a:dear(i).s - 1);
        [ds, de, dm] = regexpi(w, dateRx, 'start', 'end', 'match');
        if isempty(ds), continue; end
        pos = a - 1 + ds(end);  pe = a - 1 + de(end);
        R.value = dm{end};
        R.evidence = sprintf('[p.%s] "%s"', pageStr(pageOf(corpus, pos)), quoteAround(corpus, pos, pe, 60, 60, cfg));
        R.page = pagesToStr(pageOf(corpus, pos));  R.section = sectionAt(corpus, pos);
        R.review = true;
        R.note = 'Date taken from the FDA clearance-letter header (letter date; verify it equals the decision date).';
        R.status = assignStatus(true, false, false, false);
        return;
    end
    prep = findAll(corpus, ['Date\s+(?:Prepared|of\s+(?:Preparation|Summary))\s*:?\s*(' dateRx ')']);
    R.review = true;
    R.note = 'No clearance/decision date found.';
    if ~isempty(prep)
        R.note = [R.note ' Only a summary-preparation date appears (' prep(1).tok{1} ', p.' pageStr(prep(1).page) '), which is NOT the clearance date.'];
    end
end

function [R, info] = ruleModality(corpus, cfg)
    R = newResult('A5', 'A. Product identification', 'Imaging modality');
    info = struct('primary', '', 'all', {{}});
    n = size(cfg.modalities, 1);
    counts = zeros(1, n);  H = cell(1, n);  inIFU = false(1, n);
    for k = 1:n
        H{k} = findAll(corpus, cfg.modalities{k,2});
        counts(k) = numel(H{k});
        inIFU(k) = any(strcmp({H{k}.section}, 'IFU'));
    end
    if max(counts) == 0
        R.review = true;  R.note = 'No modality keywords found.';  return;
    end
    [mx, kp] = max(counts);
    incl = find(counts >= max(2, 0.25 * mx) | inIFU);
    [~, o] = sort(counts(incl), 'descend');  incl = incl(o);
    parts = arrayfun(@(k) sprintf('%s (%d mentions)', cfg.modalities{k,1}, counts(k)), incl, 'UniformOutput', false);
    R.value = joinList(parts, '; ');
    h = H{kp};
    pref = h(ismember({h.section}, {'IFU', 'DEVDESC'}));
    if isempty(pref), pref = h; end
    R = fillFromHits(R, pref(1), {cfg.modalities{kp,1}}, corpus, cfg);
    R.value = joinList(parts, '; ');
    R.status = assignStatus(true, any(inIFU(incl)), false, false);
    R.review = numel(incl) > 1 || ~any(inIFU(incl));
    if numel(incl) > 1, R.note = 'Multiple modalities mentioned; confirm which is the intended modality.'; end
    if ~any(inIFU(incl)), R.note = strtrim([R.note ' Modality inferred from keyword counts outside a detected indications section (keyword-based, not an explicit statement).']); end
    info.primary = cfg.modalities{kp,1};
    info.all = cfg.modalities(incl, 1)';
end

function R = ruleIntendedUse(corpus, cfg)
    R = newResult('A6', 'A. Product identification', 'Intended use / indications for use');
    G = corpus.text;  secs = corpus.sections;
    for i = 1:numel(secs)
        if ~strcmp(secs(i).name, 'IFU'), continue; end
        raw = G(secs(i).start:secs(i).stop);
        k = strfind(raw, secs(i).heading);
        if ~isempty(k), raw(k(1):k(1) + numel(secs(i).heading) - 1) = []; end
        body = strtrim(regexprep(raw, '\s+', ' '));
        body = regexprep(body, '^[:\-\s]+', '');
        if numel(body) < 40, continue; end
        truncated = numel(body) > cfg.maxIFUChars;
        if truncated, body = [body(1:cfg.maxIFUChars) ' ...']; end
        pg = pageOf(corpus, secs(i).start);
        R.value = body;
        R.evidence = sprintf('[p.%s] Section "%s": "%s"', pageStr(pg), secs(i).heading, body(1:min(end, 300)));
        R.page = pagesToStr(pg);  R.section = 'IFU';
        R.status = assignStatus(true, true, false, false);
        R.review = true;
        R.note = 'Section extent set by heading heuristics; confirm that the full indication is captured and nothing extra is included.';
        if truncated, R.note = [R.note ' Value truncated.']; end
        return;
    end
    hits = findAll(corpus, '\b(?:is|are)\s+(?:indicated|intended)\s+(?:for|to|as)\b[^.]{20,600}\.');
    if ~isempty(hits)
        h = hits(1);
        pre = corpus.text(max(1, h.s - 150):h.s - 1);
        pre = regexprep(pre, '^.*\.\s+', '');
        R.value = strtrim(regexprep([pre h.match], '\s+', ' '));
        R.evidence = sprintf('[p.%s] "%s"', pageStr(h.page), R.value(1:min(end, 300)));
        R.page = pagesToStr(h.page);  R.section = h.section;
        R.status = assignStatus(true, true, false, false);
        R.review = true;  R.note = 'No "Indications for Use" heading detected; sentence found by phrase match.';
    else
        R.review = true;  R.note = 'No indications-for-use section or sentence found.';
    end
end

%% ======================================================================
%%  SCANNER MENTIONS (shared by B, C, E)
%% ======================================================================
function M = findScannerMentions(corpus, cfg)
% Finds "<manufacturer> <model list>" phrases. Every mention keeps its position, page,
% section, quotation and a CONTEXT LABEL. A scanner named in the device description is
% NOT treated as evidence that it was in the validation data.
    G = corpus.text;
    M = struct('mfr', {}, 'model', {}, 's', {}, 'e', {}, 'stop', {}, 'page', {}, 'section', {}, ...
               'quote', {}, 'ctx', {}, 'window', {}, 'support', {});
    allH = struct('mfr', {}, 's', {}, 'e', {});
    for k = 1:size(cfg.mfrPatterns, 1)
        [s, e] = regexp(G, ['\b(?:' cfg.mfrPatterns{k,2} ')\b'], 'start', 'end', 'ignorecase');
        for i = 1:numel(s)
            allH(end+1) = struct('mfr', cfg.mfrPatterns{k,1}, 's', s(i), 'e', e(i)); %#ok<AGROW>
        end
    end
    if isempty(allH), return; end
    [~, ix] = sort([allH.s]);  allH = allH(ix);

    for i = 1:numel(allH)
        s = allH(i).s;  e = allH(i).e;
        % skip company names (Inc/GmbH...) and submitter/contact lines
        after = G(e+1:min(end, e+14));
        if ~isempty(regexpi(after, '^[\s,]*(?:GmbH|Inc\b|Corp|Ltd|LLC|AG\b|Co\b)', 'once')), continue; end
        ls = find(G(1:s) == char(10), 1, 'last'); if isempty(ls), ls = 0; end
        if ~isempty(regexpi(G(ls+1:s-1), 'Submitter|Applicant|Sponsor|Prepared\s+by|Contact|Establishment|Registration|Address', 'once')), continue; end
        if i < numel(allH), stopMax = allH(i+1).s - 1; else, stopMax = numel(G); end
        stop = min([stopMax, e + 250, numel(G)]);
        span = G(e+1:stop);
        cut = regexp(span, '\.\s+(?=[A-Z])|\n\s*\n', 'once');
        if ~isempty(cut), span = span(1:cut); stop = e + cut; end
        span = regexprep(span, '^[\s:\-\(]+', '');
        parts = regexp(span, '\s*(?:,|;|\band\b|\bor\b|\(|\)|\n|[ \t]{3,})\s*', 'split');
        nAdded = 0;
        for p = 1:numel(parts)
            part = strtrim(parts{p});
            if isempty(part), continue; end
            mdl = cleanModel(part, cfg);
            if isempty(mdl), break; end
            if nAdded > 0 && ~isempty(regexp(mdl(1), '[^A-Za-z]', 'once')) && isempty(regexp(mdl, '[A-Za-z]{2,}', 'once')), break; end
            M(end+1) = makeMention(corpus, cfg, allH(i).mfr, mdl, s, e, stop); %#ok<AGROW>
            nAdded = nAdded + 1;
            if nAdded >= 8, break; end
        end
        if nAdded == 0
            M(end+1) = makeMention(corpus, cfg, allH(i).mfr, '', s, e, e); %#ok<AGROW>
        end
    end
end

function m = makeMention(corpus, cfg, mfr, mdl, s, e, stop)
    G = corpus.text;
    win = G(max(1, s - 250):min(numel(G), stop + 250));
    sec = sectionAt(corpus, s);
    m = struct('mfr', mfr, 'model', mdl, 's', s, 'e', e, 'stop', stop, 'page', pageOf(corpus, s), ...
        'section', sec, 'quote', quoteAround(corpus, s, stop, 40, 60, cfg), ...
        'ctx', classifyScannerContext(sec, win, cfg), 'window', win, ...
        'support', ~isempty(regexpi(win, ['compatible|compatibility|\bsupported\b|\bsupports?\b|interoperab|designed\s+for\s+use\s+with|' ...
                                         'can\s+be\s+(?:used|run|applied)|intended\s+for\s+use\s+with|works?\s+with|validated\s+(?:on|with|for)'], 'once')));
end

function mdl = cleanModel(part, cfg)
% Turns a text fragment into a candidate model name (up to 5 tokens, cut at stop-words).
    mdl = '';
    part = regexprep(part, '^[\s:\-]+', '');
    toks = regexp(part, '[A-Za-z0-9][A-Za-z0-9\.\-\+]*', 'match');
    keep = {};
    for i = 1:min(numel(toks), 5)
        t = regexprep(toks{i}, '[\.\-]+$', '');
        if isempty(t) || any(strcmp(lower(t), cfg.modelStop)), break; end
        keep{end+1} = t; %#ok<AGROW>
    end
    if isempty(keep), return; end
    cand = strjoin(keep, ' ');
    if ~any(isletter(cand)), return; end
    if ~isempty(regexp(keep{1}, '^\d+$', 'once')), return; end
    if ~isempty(regexpi(keep{1}, '^\d+-?(?:slice|row|detector)', 'once')), return; end
    if isempty(regexp(cand, '[A-Z0-9]', 'once')), return; end
    mdl = cand;
end

function c = classifyScannerContext(sec, win, cfg)
    if strcmp(sec, 'PERFORMANCE')
        c = cfg.ctx.val;
    elseif ~isempty(regexpi(win, '\btrain(?:ing|ed)\b', 'once')) && isempty(regexpi(win, 'test(?:ing)?\s+(?:set|data|dataset|cases)|validation', 'once'))
        c = cfg.ctx.train;
    elseif ~isempty(regexpi(win, 'test(?:ing)?\s+(?:set|data|dataset|cases)|validation|retrospective|test\s+dataset', 'once'))
        c = cfg.ctx.near;
    else
        c = cfg.ctx.desc;
    end
end

%% ======================================================================
%%  B. SCANNER SPECIFICATIONS
%% ======================================================================
function [B1, B2] = ruleScannerSpecs(corpus, cfg, M)
    B1 = newResult('B1', 'B. Scanner specifications', 'Scanner manufacturer(s) listed in document');
    B2 = newResult('B2', 'B. Scanner specifications', 'Exact scanner model(s) listed');
    if isempty(M)
        B1.review = true;  B1.note = 'No manufacturer from the built-in list (cfg.mfrPatterns) was found; add vendors if needed.';
        B2.review = true;  B2.note = B1.note;  return;
    end
    [umf, ia] = unique({M.mfr}, 'stable');
    B1 = fillFromHits(B1, M(ia), umf, corpus, cfg);
    B1.status = assignStatus(true, true, false, false);
    B1.review = true;
    B1.note = 'Mentions anywhere in the document; does not imply use in validation (see C5/C6).';
    idx = find(~cellfun(@isempty, {M.model}));
    if isempty(idx)
        B2 = fillFromHits(B2, M(ia(1)), {''}, corpus, cfg);
        B2.value = '';
        B2.status = assignStatus(true, false, false, false);
        B2.review = true;  B2.note = 'Manufacturer(s) named but no model name could be parsed.';
    else
        Mi = M(idx);
        keys = cellfun(@(a, b) [a ' ' b], {Mi.mfr}, {Mi.model}, 'UniformOutput', false);
        [uk, ik] = unique(keys, 'stable');
        B2 = fillFromHits(B2, Mi(ik), uk, corpus, cfg);
        B2.status = assignStatus(true, true, false, false);
        B2.review = true;
        B2.note = 'Model names are parsed heuristically (text after the manufacturer name); verify spelling and completeness against the PDF.';
    end
end

function R = ruleParam(corpus, cfg, id, field, pattern, kind)
% Generic reconstruction-parameter rule (algorithm / kernel / strength / slice).
    R = newResult(id, 'B. Scanner specifications', field);
    hits = findAll(corpus, pattern);
    if isempty(hits), return; end
    vals = cellfun(@(m) strtrim(regexprep(m, '\s+', ' ')), {hits.match}, 'UniformOutput', false);
    [hits, vals] = dedupHits(hits, vals);
    R = fillFromHits(R, hits, vals, corpus, cfg);
    complete = true;
    switch kind
        case 'algorithm'
            complete = any(~cellfun(@isempty, regexpi(vals, cfg.paramPat.algorithmSpecific, 'once')));
            if ~complete, R.note = 'Only a generic algorithm class is named (no product/vendor-specific algorithm).'; end
        case 'kernel'
            complete = any(~cellfun(@isempty, regexpi(vals, cfg.paramPat.kernelName, 'once')));
            if ~complete, R.note = 'The word "kernel" appears but no identifiable kernel name/code was detected.'; end
            R.review = true;
        case 'strength'
            R.review = true;
        case 'slice'
            R.review = numel(vals) > 1;
    end
    if any(strcmp(kind, {'algorithm','strength','slice'})) && numel(vals) > 1
        R.note = strtrim([R.note ' Multiple values; they may belong to different scanners or datasets (not linked here).']);
    end
    R.status = assignStatus(true, complete, false, false);
end

function R = ruleAcquisition(corpus, cfg)
    R = newResult('B7', 'B. Scanner specifications', 'Other acquisition parameters / protocols');
    allHits = struct('s', {}, 'e', {}, 'match', {}, 'tok', {}, 'page', {}, 'section', {});
    allVals = {};
    for k = 1:size(cfg.acqPatterns, 1)
        h = findAll(corpus, cfg.acqPatterns{k,2});
        if isempty(h), continue; end
        v = cellfun(@(m) [cfg.acqPatterns{k,1} ': ' strtrim(regexprep(m, '\s+', ' '))], {h.match}, 'UniformOutput', false);
        [h, v] = dedupHits(h, v);
        h = h(1:min(2, numel(h)));  v = v(1:min(2, numel(v)));    % keep first two per parameter type
        allHits = [allHits, h]; %#ok<AGROW>
        allVals = [allVals, v]; %#ok<AGROW>
    end
    if isempty(allHits), return; end
    R = fillFromHits(R, allHits, allVals, corpus, cfg);
    R.status = assignStatus(true, true, false, false);
    R.review = true;
    R.note = 'Pattern-based; values are not linked to specific scanners. Only the first two matches per parameter type are listed.';
end

%% ======================================================================
%%  C. VALIDATION DATASET
%% ======================================================================
function R = ruleCountField(corpus, cfg, id, fieldName, patterns, exclRx)
% Counts are kept per field and unit (patients / studies / images / sites are never merged).
    R = newResult(id, 'C. Validation dataset', fieldName);
    hits = struct('s', {}, 'e', {}, 'match', {}, 'tok', {}, 'page', {}, 'section', {}, 'unit', {});
    for p = 1:size(patterns, 1)
        h = findAll(corpus, patterns{p,1});
        for i = 1:numel(h)
            if ~isempty(regexpi(h(i).match, exclRx, 'once')), continue; end
            h(i).unit = patterns{p,2};                       % default unit label for this pattern
            if numel(h(i).tok) >= 2, h(i).unit = lower(regexprep(h(i).tok{2}, '\s+', ' ')); end
            hits(end+1) = h(i); %#ok<AGROW>
        end
    end
    if isempty(hits), return; end
    [hits, fallback] = preferSection(hits, {'PERFORMANCE'});
    units = {hits.unit};
    nums = arrayfun(@(i) parseCountToken(hits(i).tok{1}), 1:numel(hits));
    vals = arrayfun(@(i) sprintf('%s %s', fmtNum(nums(i)), strtrim(units{i})), 1:numel(hits), 'UniformOutput', false);
    ok = ~isnan(nums) & nums > 0;
    hits = hits(ok);  vals = vals(ok);  nums = nums(ok);
    if isempty(hits), return; end
    [~, ia] = unique(vals, 'stable');             % dedupe identical statements
    hits = hits(ia);  vals = vals(ia);  nums = nums(ia);
    R.details = makeDetails(fieldName, hits, vals, corpus, cfg);
    hasTotal = false(1, numel(hits));
    for i = 1:numel(hits)
        hasTotal(i) = ~isempty(regexpi(quoteAround(corpus, hits(i).s, hits(i).e, 45, 0, cfg), 'total|overall|in\s+all|combined|altogether', 'once'));
    end
    if any(hasTotal) && numel(unique(nums(hasTotal))) == 1
        k = find(hasTotal, 1);
        R = fillFromHits(R, hits(k), vals(k), corpus, cfg);
        R.status = assignStatus(true, true, false, false);
        if numel(unique(nums)) > 1
            R.review = true;  R.note = ['Other counts also reported: ' joinList(vals, '; ')];
        end
    elseif numel(unique(nums)) == 1
        R = fillFromHits(R, hits, vals, corpus, cfg);
        R.status = assignStatus(true, true, false, false);
    else
        R = fillFromHits(R, hits, vals, corpus, cfg);
        R.status = assignStatus(true, false, true, false);
        R.review = true;
        R.note = 'Multiple distinct counts; they may describe different datasets/subgroups (e.g., training vs test) and no explicit total was identified.';
    end
    if fallback
        R.review = true;  R.note = strtrim([R.note ' Found outside a detected performance/clinical section.']);
    end
    if any(strcmp(id, {'C2'}))
        R.note = strtrim([R.note ' "Cases" may mean patients or studies; unit is reported as written.']);
        R.review = true;
    end
end

function n = parseCountToken(t)
    t = lower(strtrim(t));
    words = {'one','two','three','four','five','six','seven','eight','nine','ten','eleven','twelve','thirteen','fourteen','fifteen'};
    k = find(strcmp(t, words), 1);
    if ~isempty(k), n = k; else, n = str2double(strrep(t, ',', '')); end
end

function s = fmtNum(n)
    s = regexprep(sprintf('%d', round(n)), '(\d)(?=(\d{3})+$)', '$1,');
end

function R = ruleSiteSplits(corpus, cfg)
    R = newResult('C4', 'C. Validation dataset', 'Sites used for training / validation / testing');
    splitW = '\b(?:training|development|tuning|validation|test(?:ing)?|hold-?out)\b';
    siteN  = '(?:sites?|institutions?|hospitals?|centers?|centres?|facilities)';
    p1 = ['\b' cfg.nw '\s+(?:[A-Za-z\-]+\s+){0,3}?' siteN '\b[^.\r\n]{0,100}?' splitW];
    p2 = [splitW '[^.\r\n]{0,120}?\b' cfg.nw '\s+(?:[A-Za-z\-]+\s+){0,3}?' siteN '\b'];
    hits = [findAll(corpus, p1), findAll(corpus, p2)];
    if isempty(hits), return; end
    [~, o] = sort([hits.s]);  hits = hits(o);
    vals = cellfun(@(m) strtrim(regexprep(m, '\s+', ' ')), {hits.match}, 'UniformOutput', false);
    [hits, vals] = dedupHits(hits, vals);
    R = fillFromHits(R, hits, vals, corpus, cfg);
    splitsSeen = {};
    for i = 1:numel(vals)
        splitsSeen = [splitsSeen, lower(regexpi(vals{i}, splitW, 'match'))]; %#ok<AGROW>
    end
    splitsSeen = regexprep(unique(splitsSeen), 'testing', 'test');
    splitsSeen = unique(regexprep(splitsSeen, '^(development|tuning)$', 'training'));
    R.status = assignStatus(true, numel(splitsSeen) >= 2, false, false);
    R.review = true;
    R.note = 'Phrase-level match of site counts near split words; confirm which split each count refers to.';
end

function [C5, C6, C7] = ruleValidationScanners(corpus, cfg, M)
    C5 = newResult('C5', 'C. Validation dataset', 'Scanner manufacturers represented in validation data');
    C6 = newResult('C6', 'C. Validation dataset', 'Scanner models represented in validation data');
    C7 = newResult('C7', 'C. Validation dataset', 'Sample size per scanner/model');
    note0 = ['Only scanner mentions inside a performance/validation section or near validation/test-set wording are considered. ' ...
             'Scanners named only in the device description are NOT assumed to be in the validation data.'];
    if isempty(M)
        C5.review = true; C5.note = 'No scanner mentions found.'; C6.review = true; C6.note = C5.note; C7.review = true; C7.note = C5.note; return;
    end
    isVal  = strcmp({M.ctx}, cfg.ctx.val);
    isNear = strcmp({M.ctx}, cfg.ctx.near);
    Mv = M(isVal | isNear);
    if isempty(Mv)
        C5.note = ['Scanners are named elsewhere in the document but not in validation context. ' note0];
        C5.review = true;  C6.note = C5.note;  C6.review = true;  C7.note = C5.note;  C7.review = true;
        return;
    end
    anyVal = any(strcmp({Mv.ctx}, cfg.ctx.val));
    [umf, ia] = unique({Mv.mfr}, 'stable');
    C5 = fillFromHits(C5, Mv(ia), umf, corpus, cfg);
    C5.status = assignStatus(true, anyVal, false, false);
    C5.review = true;  C5.note = note0;
    idx = find(~cellfun(@isempty, {Mv.model}));
    if ~isempty(idx)
        Mm = Mv(idx);
        keys = cellfun(@(a, b) [a ' ' b], {Mm.mfr}, {Mm.model}, 'UniformOutput', false);
        [uk, ik] = unique(keys, 'stable');
        C6 = fillFromHits(C6, Mm(ik), uk, corpus, cfg);
        C6.status = assignStatus(true, any(strcmp({Mm.ctx}, cfg.ctx.val)), false, false);
        C6.review = true;  C6.note = note0;
        % per-scanner sample sizes (same passage only)
        vals = {};  hh = Mm([]);
        for i = 1:numel(Mm)
            win = corpus.text(max(1, Mm(i).s - 100):min(end, Mm(i).stop + 250));
            nn = regexpi(win, '\bn\s*=\s*\d[\d,]*|\b\d[\d,]*\s+(?:patients?|studies|cases|scans?|exams?|images?)\b', 'match');
            if ~isempty(nn)
                vals{end+1} = sprintf('%s %s: %s', Mm(i).mfr, Mm(i).model, strtrim(nn{1})); %#ok<AGROW>
                hh(end+1) = Mm(i); %#ok<AGROW>
            end
        end
        if ~isempty(hh)
            C7 = fillFromHits(C7, hh, vals, corpus, cfg);
            C7.details = makeDetails(C7.field, hh, vals, corpus, cfg);
            C7.status = assignStatus(true, false, true, false);
            C7.review = true;
            C7.note = ['A count appears in the same passage as the scanner name; the link between count and scanner is NOT confirmed. ' ...
                       'Bare numbers in table cells (no n= or unit) are not detected.'];
        end
    else
        C6.status = assignStatus(true, false, false, false);
        C6.review = true;  C6.note = ['Manufacturer(s) named in validation context but no model parsed. ' note0];
    end
end

%% ======================================================================
%%  D. VALIDATION DESIGN & PERFORMANCE
%% ======================================================================
function [R, info] = ruleIndependence(corpus, cfg)
    R = newResult('D2', 'D. Validation design and performance', 'Independent institutions/sites used');
    siteN = '(?:sites?|institutions?|hospitals?|centers?|centres?|facilities|health\s+systems?)';
    pSite = ['\b(?:independent|separate|distinct|different|unseen|external|non-?overlapping)\b[^.\r\n]{0,40}?\b' siteN '\b|' ...
             '\b' siteN '\b[^.\r\n]{0,60}?\b(?:not\s+(?:used|included|part)|different\s+from|independent\s+(?:of|from)|separate\s+from)\b[^.\r\n]{0,40}\b(?:training|development)'];
    pData = ['\bindependent\s+(?:of|from)\s+(?:the\s+)?(?:training|development)\b|' ...
             '\b(?:not|never)\s+(?:been\s+)?used\s+(?:for|in|during)\s+(?:the\s+)?(?:training|development|model\s+development)\b|' ...
             '\bno\s+(?:patient\s+)?overlap\b|\bindependent\s+(?:test|validation)\s+(?:set|data|dataset|cohort)\b'];
    hs = dropNegated(findAll(corpus, pSite), corpus);
    hd = dropNegated(findAll(corpus, pData), corpus);
    info = struct('anySite', ~isempty(hs), 'anyData', ~isempty(hd));
    if ~isempty(hs)
        vals = cellfun(@(m) strtrim(regexprep(m, '\s+', ' ')), {hs.match}, 'UniformOutput', false);
        R = fillFromHits(R, hs, vals, corpus, cfg);
        R.status = assignStatus(true, true, false, false);
        R.review = true;  R.note = 'Phrase match; confirm the statement refers to the evaluation (test/validation) sites.';
    elseif ~isempty(hd)
        vals = cellfun(@(m) strtrim(regexprep(m, '\s+', ' ')), {hd.match}, 'UniformOutput', false);
        R = fillFromHits(R, hd, vals, corpus, cfg);
        R.status = assignStatus(true, false, false, false);
        R.review = true;  R.note = 'Data independence from training is stated, but independence at the institution/site level is not.';
    end
end

function R = ruleInternalExternal(corpus, cfg, indep)
    R = newResult('D1', 'D. Validation design and performance', 'Internal vs external evaluation');
    ext = dropNegated(findAll(corpus, ['\bexternal(?:ly)?\b[^.\r\n]{0,40}?\b(?:validation|validated|test|testing|tested|dataset|data|set|sites?|institutions?|evaluation|evaluated|cohort)\b|' ...
                                       '\b(?:validation|testing|test)\b[^.\r\n]{0,30}?\bexternal\b']), corpus);
    intH = dropNegated(findAll(corpus, ['\binternal(?:ly)?\b[^.\r\n]{0,40}?\b(?:validation|validated|test|testing|tested|dataset|data|set|evaluation|evaluated|cohort)\b|' ...
                                        '\bhold-?out\b|\bheld[- ]out\b|\brandomly\s+split\b|\bcross-?validation\b']), corpus);
    hasE = ~isempty(ext);  hasI = ~isempty(intH);
    hits = struct('s', {}, 'e', {}, 'match', {}, 'tok', {}, 'page', {}, 'section', {});  vals = {};
    if hasE, hits = [hits, ext(1)]; vals{end+1} = ['External: ' strtrim(regexprep(ext(1).match, '\s+', ' '))]; end
    if hasI, hits = [hits, intH(1)]; vals{end+1} = ['Internal/hold-out: ' strtrim(regexprep(intH(1).match, '\s+', ' '))]; end
    if hasE || hasI
        R = fillFromHits(R, hits, vals, corpus, cfg);
        R.review = true;
        if hasE && ~hasI
            R.value = 'External evaluation language found';
            R.status = assignStatus(true, true, false, false);
        elseif hasI && ~hasE
            R.value = 'Internal / hold-out evaluation language found (no explicit external-validation wording)';
            R.status = assignStatus(true, true, false, false);
        else
            R.value = 'Both internal/hold-out and external wording found';
            R.status = assignStatus(true, true, false, false);
            R.note = 'Both kinds of wording present; determine manually which dataset each describes.';
        end
        R.note = strtrim([R.note ' Keyword-based.']);
    elseif indep.anySite || indep.anyData
        R.value = 'Independence from training data is stated but the evaluation is not labelled internal/external';
        R.status = assignStatus(true, false, false, false);
        R.review = true;  R.note = 'See D2 for the supporting text.';
    end
end

function [D3, D4, D5, rows] = rulePerformance(corpus, cfg)
    D3 = newResult('D3', 'D. Validation design and performance', 'Performance metrics reported');
    D4 = newResult('D4', 'D. Validation design and performance', 'Performance reported by scanner manufacturer/model vs aggregate');
    D5 = newResult('D5', 'D. Validation design and performance', 'Scanner-specific performance values (with n and CI)');
    rows = struct('metric', {}, 'value', {}, 'ci', {}, 'n', {}, 'level', {}, 'subgroup', {}, 'page', {}, ...
                  'section', {}, 'quote', {}, 'review', {}, 'note', {});
    nM = size(cfg.metrics, 1);
    H = cell(1, nM);
    for k = 1:nM
        H{k} = findAll(corpus, cfg.metrics{k,2});
    end
    present = find(cellfun(@numel, H) > 0);
    if isempty(present)
        D3.review = true;  D3.note = 'No metric names found.';  D4.review = true;  D5.review = true;  return;
    end
    mfrRx = strjoin(cfg.mfrPatterns(:,2)', '|');
    scanRx = ['\b(?:' mfrRx ')\b|scanner\s+(?:model|manufacturer|vendor|type)|(?:by|per|for\s+each)\s+(?:scanner|vendor|manufacturer)'];
    subRx  = 'sub-?group|stratif|by\s+(?:age|sex|gender|race|ethnic\w*|site|slice|kernel|lesion|size)|per[\s-]site';
    ciRx1  = '(?:CI|confidence\s+interval)[^\d\r\n]{0,15}(\d{1,3}(?:\.\d+)?)\s*%?\s*(?:-|to|,|;)\s*(\d{1,3}(?:\.\d+)?)\s*%?';
    ciRx2  = '[\[\(]\s*(\d{1,3}(?:\.\d+)?)\s*%?\s*(?:-|to|,)\s*(\d{1,3}(?:\.\d+)?)\s*%?\s*[\]\)]';
    seen = {};
    metricsWithValues = {};
    for k = present
        [hk, fb] = preferSection(H{k}, {'PERFORMANCE'}); %#ok<ASGLU>
        for i = 1:numel(hk)
            if numel(rows) >= 100, break; end
            h = hk(i);
            win = corpus.text(h.s:min(numel(corpus.text), h.e + 200));
            c = regexp(win, '\.\s+[A-Z]|\n\s*\n', 'once');
            if ~isempty(c), win = win(1:c); end
            tailStart = numel(h.match) + 1;
            if tailStart <= numel(win)
                cutAt = Inf;
                for k2 = present
                    if k2 == k, continue; end
                    kk = regexpi(win(tailStart:end), cfg.metrics{k2,2}, 'once');
                    if ~isempty(kk), cutAt = min(cutAt, tailStart + kk - 2); end
                end
                if isfinite(cutAt), win = win(1:cutAt); end
            end
            tailTxt = win(min(tailStart, numel(win)+1):end);
            % confidence interval (kept in original wording)
            ci = '';  rest = tailTxt;
            for rx = {ciRx1, ciRx2}
                [st, en] = regexpi(rest, rx{1}, 'start', 'end');
                if ~isempty(st)
                    ci = strtrim(rest(st(1):en(1)));
                    rest(st(1):en(1)) = ' ';  break;
                end
            end
            val = regexp(rest, '(?<![\d.])\d{1,3}(?:\.\d+)?\s*%', 'match', 'once');
            if isempty(val), val = regexp(rest, '(?<![\d.])(?:0?\.\d{2,3}|1\.0{1,3})(?![\d%])', 'match', 'once'); end
            nn = regexp(win, '\bn\s*=\s*\d[\d,]*|\b\d[\d,]*\s+(?:patients?|studies|cases|scans?|exams?|images?|subjects?)\b', 'match', 'once');
            if isempty(val) && isempty(ci), continue; end
            lvl = 'Aggregate/unspecified (verify)';  sub = '';  rv = false;  nt = '';
            ms = regexpi(win, scanRx, 'match', 'once');
            if ~isempty(ms)
                lvl = 'Scanner-specific (candidate)';  sub = strtrim(regexprep(ms, '\s+', ' '));  rv = true;
                nt = 'Scanner wording appears in the same passage; confirm that the value is stratified by scanner.';
            else
                mg = regexpi(win, subRx, 'match', 'once');
                if ~isempty(mg), lvl = 'Other subgroup (candidate)'; sub = strtrim(mg); rv = true; nt = 'Subgroup wording in passage; confirm.'; end
            end
            key = lower([cfg.metrics{k,1} '|' val '|' ci '|' pageStr(h.page)]);
            if any(strcmp(seen, key)), continue; end
            seen{end+1} = key; %#ok<AGROW>
            rows(end+1) = struct('metric', cfg.metrics{k,1}, 'value', val, 'ci', ci, 'n', strtrim(nn), 'level', lvl, ...
                'subgroup', sub, 'page', pageStr(h.page), 'section', h.section, ...
                'quote', quoteAround(corpus, h.s, h.s + numel(win) - 1, 20, 0, cfg), ...
                'review', rv || isempty(val), 'note', nt); %#ok<AGROW>
            metricsWithValues{end+1} = cfg.metrics{k,1}; %#ok<AGROW>
        end
    end
    % D3 summary
    labels = cfg.metrics(present, 1)';
    vals = cellfun(@(m) m, labels, 'UniformOutput', false);
    hd = cellfun(@(h) h(1), H(present), 'UniformOutput', false);
    hd = [hd{:}];
    D3 = fillFromHits(D3, hd, vals, corpus, cfg);
    D3.status = assignStatus(true, ~isempty(rows), false, false);
    D3.review = true;
    D3.note = sprintf('%d numeric result row(s) parsed (see Performance Results).', numel(rows));
    if isempty(rows)
        D3.note = 'Metric names appear but no numeric values were parsed from the text (results may be in tables/figures stored as images); check manually.';
    end
    % D4 / D5
    scanRows = rows(strcmp({rows.level}, 'Scanner-specific (candidate)'));
    stratText = findAll(corpus, '(?:by|across|per|stratified\s+by|for\s+each)\s+(?:CT\s+)?(?:scanner|vendor|manufacturer|scanner\s+model|scanner\s+manufacturer)');
    stratText = dropNegated(stratText, corpus);
    if ~isempty(scanRows) || ~isempty(stratText)
        D4.value = 'Scanner-stratified reporting appears to be present';
        if ~isempty(stratText)
            D4.evidence = sprintf('[p.%s] "%s"', pageStr(stratText(1).page), quoteAround(corpus, stratText(1).s, stratText(1).e, 60, 100, cfg));
            D4.page = pagesToStr(stratText(1).page);  D4.section = stratText(1).section;
        else
            D4.evidence = sprintf('[p.%s] "%s"', scanRows(1).page, scanRows(1).quote);  D4.page = scanRows(1).page;  D4.section = scanRows(1).section;
        end
        D4.status = assignStatus(true, true, false, false);
        D4.review = true;  D4.note = 'Wording-based; verify against the tables.';
    else
        D4.value = 'Aggregate only (apparent)';
        D4.status = assignStatus(true, false, false, false);
        D4.review = true;
        D4.note = ['No scanner-stratified performance wording found in extracted text. Absence in text cannot be proven ' ...
                   '(image-based tables/figures are not read), hence "Partially disclosed" and manual review.'];
        D4.evidence = D3.evidence;  D4.page = D3.page;  D4.section = D3.section;
    end
    if ~isempty(scanRows)
        D5.value = sprintf('%d scanner-labelled result row(s) found (see Performance Results)', numel(scanRows));
        D5.evidence = sprintf('[p.%s] "%s"', scanRows(1).page, scanRows(1).quote);
        D5.page = strjoin(unique({scanRows.page}), ', ');  D5.section = scanRows(1).section;
        full = any(~cellfun(@isempty, {scanRows.value}) & ~cellfun(@isempty, {scanRows.ci}) & ~cellfun(@isempty, {scanRows.n}));
        D5.status = assignStatus(true, full, false, false);
        D5.review = true;
        D5.note = 'Scanner linkage of each row is heuristic; check value, n and CI against the source table.';
    end
end

%% ======================================================================
%%  E. INPUT REQUIREMENTS AND LIMITATIONS
%% ======================================================================
function R = ruleInputType(corpus, cfg)
    R = newResult('E1', 'E. Input requirements and limitations', 'Required image type / series');
    pat = ['\b(?:input|inputs|required|requires?|requirements?)\b[^.\r\n]{0,150}\b(?:DICOM|series|axial|non-?contrast|contrast|images?|sequence|view|projection)\b[^.\r\n]{0,100}'];
    hits = findAll(corpus, pat);
    if isempty(hits), return; end
    [hits, fb] = preferSection(hits, {'IFU', 'DEVDESC', 'TECH', 'LIMITATIONS'});
    vals = cellfun(@(m) strtrim(regexprep(m, '\s+', ' ')), {hits.match}, 'UniformOutput', false);
    [hits, vals] = dedupHits(hits, vals);
    R = fillFromHits(R, hits(1:min(3, numel(hits))), vals(1:min(3, numel(vals))), corpus, cfg);
    R.status = assignStatus(true, true, false, false);
    R.review = true;  R.note = 'Sentence-level phrase match; confirm it states a requirement.';
    if fb, R.note = [R.note ' Found outside indications/description sections.']; end
end

function R = ruleAcqRequirements(corpus, cfg)
    R = newResult('E2', 'E. Input requirements and limitations', 'Image acquisition / reconstruction requirements');
    terms = 'slice\s+thickness|kernel|reconstruction|kVp|field\s+of\s+view|pitch|collimation|contrast|acquisition|field\s+strength';
    req   = 'must|required|requires?|should|compatible|supported|accepted|between|at\s+least|no\s+more\s+than|less\s+than|greater\s+than|<=|>=|<|>';
    pat = ['\b(?:' terms ')\b[^.\r\n]{0,120}?(?:\b(?:' req ')\b|<=|>=|<|>)|\b(?:must|required|requires?)\b[^.\r\n]{0,120}?\b(?:' terms ')\b'];
    hits = findAll(corpus, pat);
    if isempty(hits), return; end
    [hits, fb] = preferSection(hits, {'IFU', 'DEVDESC', 'TECH', 'LIMITATIONS'});
    vals = cellfun(@(m) strtrim(regexprep(m, '\s+', ' ')), {hits.match}, 'UniformOutput', false);
    [hits, vals] = dedupHits(hits, vals);
    R = fillFromHits(R, hits(1:min(3, numel(hits))), vals(1:min(3, numel(vals))), corpus, cfg);
    R.status = assignStatus(true, false, false, false);
    R.review = true;  R.note = 'Keyword-proximity match; may describe the validation protocol rather than a requirement.';
    if fb, R.note = [R.note ' Found outside indications/description sections.']; end
end

function R = ruleSupportedScanners(corpus, cfg, M)
    R = newResult('E3', 'E. Input requirements and limitations', 'Explicitly supported scanner models/configurations');
    if isempty(M)
        R.review = true;  R.note = 'No scanner mentions found.';  return;
    end
    sup = M([M.support]);
    if isempty(sup)
        R.value = '';
        R.status = assignStatus(true, false, true, false);
        R.review = true;
        R.evidence = sprintf('[p.%s] "%s"', pageStr(M(1).page), M(1).quote);
        R.page = pagesToStr(M(1).page);  R.section = M(1).section;
        R.note = 'Scanners are named but no compatibility/support wording was found near them, so it is unclear whether they are the supported set.';
        return;
    end
    idx = find(~cellfun(@isempty, {sup.model}));
    if ~isempty(idx)
        Ms = sup(idx);
        keys = cellfun(@(a, b) [a ' ' b], {Ms.mfr}, {Ms.model}, 'UniformOutput', false);
        [uk, ik] = unique(keys, 'stable');
        R = fillFromHits(R, Ms(ik), uk, corpus, cfg);
        R.status = assignStatus(true, true, false, false);
    else
        [umf, ia] = unique({sup.mfr}, 'stable');
        R = fillFromHits(R, sup(ia), umf, corpus, cfg);
        R.status = assignStatus(true, false, false, false);
    end
    R.review = true;  R.note = 'Based on support/compatibility wording within 250 characters of the scanner name; confirm.';
end

function R = ruleLimitations(corpus, cfg)
    R = newResult('E4', 'E. Input requirements and limitations', 'Scanner-related limitations or exclusions');
    pat = ['\b(?:not\s+(?:intended|validated|tested|evaluated|supported|compatible|designed)|exclud(?:e|ed|es|ing)|limitations?|contraindicat\w+|' ...
           'do(?:es)?\s+not\s+(?:support|work|apply)|only\s+(?:validated|tested|compatible|supported))\b[^.\r\n]{0,200}?' ...
           '\b(?:scanners?|vendors?|manufacturers?|kernels?|slice\s+thickness|reconstruction|acquisition|protocols?|field\s+strength)\b'];
    hits = findAll(corpus, pat);
    if ~isempty(hits)
        vals = cellfun(@(m) strtrim(regexprep(m, '\s+', ' ')), {hits.match}, 'UniformOutput', false);
        [hits, vals] = dedupHits(hits, vals);
        R = fillFromHits(R, hits(1:min(3, numel(hits))), vals(1:min(3, numel(vals))), corpus, cfg);
        R.status = assignStatus(true, true, false, false);
        R.review = true;  R.note = 'Phrase match; confirm it is a scanner-related limitation.';
    elseif any(strcmp({corpus.sections.name}, 'LIMITATIONS'))
        R.review = true;
        R.note = 'A limitations/warnings section was detected but no scanner-related limitation wording was found in it.';
    end
end

%% ======================================================================
%%  POST-PROCESSING AND CROSS-CHECKS
%% ======================================================================
function results = applyNotApplicable(results, modInfo)
% N/A only when the sole detected modality is X-ray or ultrasound AND nothing was found.
    if numel(modInfo.all) ~= 1 || ~any(strcmp(modInfo.primary, {'X-ray', 'Ultrasound'})), return; end
    for id = {'B3', 'B4', 'B5', 'B6'}
        i = find(strcmp({results.id}, id{1}), 1);
        if ~isempty(i) && strcmp(results(i).status, 'Not disclosed')
            results(i).status = assignStatus(false, false, false, true);
            results(i).review = true;
            results(i).note = sprintf('Marked N/A because the only detected modality keyword is %s; confirm manually.', modInfo.primary);
        end
    end
end

function [results, msgs] = crossChecks(results, perfRows)
    msgs = {};
    idx = @(id) find(strcmp({results.id}, id), 1);
    st  = @(id) results(idx(id)).status;
    if any(strcmp({perfRows.level}, 'Scanner-specific (candidate)')) && strcmp(st('C6'), 'Not disclosed')
        msgs{end+1} = 'CHECK: scanner-labelled performance rows exist but no validation-context scanner models were identified (C6).';
        for id = {'C6', 'D5'}, results(idx(id{1})).review = true; end
    end
    if strcmp(st('D5'), 'Disclosed') && strcmp(st('C7'), 'Not disclosed')
        msgs{end+1} = 'CHECK: D5 is Disclosed but no per-scanner sample sizes were found (C7).';
        results(idx('C7')).review = true;
    end
    if any(strcmp(st('A3'), {'Unclear', 'Not disclosed'}))
        msgs{end+1} = 'CHECK: 510(k) number could not be confirmed; verify manually.';
    end
    if strcmp(st('A1'), 'Not disclosed')
        msgs{end+1} = 'CHECK: product name not found; verify manually.';
    end
    n = sum(strcmp({results.status}, 'Unclear'));
    if n > 0, msgs{end+1} = sprintf('%d field(s) have status Unclear (conflicting or ambiguous text).', n); end
end

%% ======================================================================
%%  EXCEL EXPORT
%% ======================================================================
function exportWorkbook(xlsxFile, results, perfRows, M, cfg, runInfo)
    yn = @(b) iff(b, 'Yes', 'No');

    % ---- 1. Product Summary (one row) ----
    ids = {'A1', 'A2', 'A3', 'A4', 'A5', 'A6'};
    names = {'Device_Name', 'Manufacturer', 'K_Number_510k', 'Clearance_Date', 'Imaging_Modality', 'Intended_Use'};
    vn = {};  row = {};
    for i = 1:numel(ids)
        r = results(strcmp({results.id}, ids{i}));
        vn = [vn, {names{i}, [names{i} '_Status'], [names{i} '_Page']}]; %#ok<AGROW>
        row = [row, {r.value, r.status, r.page}]; %#ok<AGROW>
    end
    flagged = results([results.review]);
    vn = [vn, {'Source_PDF', 'Extraction_Backend', 'Fields_Needing_Manual_Review'}];
    row = [row, {runInfo.pdf, runInfo.backend, joinList({flagged.id}, ', ')}];
    T1 = makeTable(row, vn);

    % ---- 2. Scanner Configurations ----
    vn2 = {'Scanner_Manufacturer', 'Scanner_Model', 'Model_Disclosure_Status', 'Context_Labels', 'Sections', 'PDF_Pages', ...
           'Mentions', 'Recon_Algorithm_SamePassage', 'Recon_Kernel_SamePassage', 'Recon_Strength_SamePassage', ...
           'Slice_Thickness_SamePassage', 'Other_Acquisition_SamePassage', 'Supporting_Quote', 'Needs_Manual_Review', 'Note'};
    rows2 = {};
    if ~isempty(M)
        keys = cellfun(@(a, b) lower([a '|' b]), {M.mfr}, {M.model}, 'UniformOutput', false);
        [uk, ~, ic] = unique(keys, 'stable');
        modelMfr = unique({M(~cellfun(@isempty, {M.model})).mfr});
        for g = 1:numel(uk)
            sel = M(ic == g);
            if isempty(sel(1).model) && any(strcmp(modelMfr, sel(1).mfr)), continue; end   % manufacturer already covered by model rows
            lp = struct('algorithm', {{}}, 'kernel', {{}}, 'strength', {{}}, 'slice', {{}}, 'acq', {{}});
            for j = 1:numel(sel)
                q = localParams(sel(j).window, cfg);
                lp.algorithm = [lp.algorithm, q.algorithm];  lp.kernel = [lp.kernel, q.kernel];
                lp.strength  = [lp.strength,  q.strength];   lp.slice  = [lp.slice,  q.slice];
                lp.acq       = [lp.acq,       q.acq];
            end
            pgs = [sel.page];
            if isempty(sel(1).model), stt = 'Partially disclosed'; mdlTxt = ''; else, stt = 'Disclosed'; mdlTxt = sel(1).model; end
            rows2(end+1, :) = {sel(1).mfr, mdlTxt, stt, joinList(uniqueStable({sel.ctx}), '; '), ...
                joinList(uniqueStable({sel.section}), '; '), pagesToStr(pgs), sprintf('%d', numel(sel)), ...
                joinList(uniqueStable(lp.algorithm), '; '), joinList(uniqueStable(lp.kernel), '; '), ...
                joinList(uniqueStable(lp.strength), '; '), joinList(uniqueStable(lp.slice), '; '), ...
                joinList(uniqueStable(lp.acq), '; '), sprintf('[p.%s] "%s"', pageStr(sel(1).page), sel(1).quote), ...
                'Yes', ['Heuristic parse. "SamePassage" columns are values found within ~250 characters of the mention and are NOT confirmed to apply to this scanner.' ...
                        iff(isempty(mdlTxt), ' Model not identified.', '')]}; %#ok<AGROW>
        end
    end
    b3 = results(strcmp({results.id}, 'B3'));  b4 = results(strcmp({results.id}, 'B4'));
    b5 = results(strcmp({results.id}, 'B5'));  b6 = results(strcmp({results.id}, 'B6'));
    b7 = results(strcmp({results.id}, 'B7'));
    rows2(end+1, :) = {'(document-level)', '(not linked to a specific scanner)', ...
        iff(isempty(M), 'Not disclosed', 'n/a'), '', '', '', '', b3.value, b4.value, b5.value, b6.value, b7.value, ...
        '', 'Yes', ['Doc-wide values from fields B3-B7 (statuses: algorithm=' b3.status ', kernel=' b4.status ...
                    ', strength=' b5.status ', slice=' b6.status ', other=' b7.status ').']};
    T2 = makeTable(rows2, vn2);

    % ---- 3. Validation Dataset ----
    vn3 = {'Row_Type', 'Field', 'Value', 'Disclosure_Status', 'PDF_Page', 'Section', 'Supporting_Evidence', 'Needs_Manual_Review', 'Note'};
    rows3 = {};
    for i = 1:numel(results)
        r = results(i);
        if ~strncmp(r.group, 'C.', 2), continue; end
        rows3(end+1, :) = {'Summary', r.field, r.value, r.status, r.page, r.section, r.evidence, yn(r.review), r.note}; %#ok<AGROW>
        for j = 1:numel(r.details)
            d = r.details(j);
            rows3(end+1, :) = {'Detail', r.field, d.value, '', d.page, d.section, d.quote, '', ''}; %#ok<AGROW>
        end
    end
    T3 = makeTable(rows3, vn3);

    % ---- 4. Performance Results ----
    vn4 = {'Metric', 'Value', 'Confidence_Interval', 'N_Near_Value', 'Reporting_Level', 'Scanner_or_Subgroup_Text', ...
           'PDF_Page', 'Section', 'Supporting_Quote', 'Needs_Manual_Review', 'Note'};
    rows4 = cell(numel(perfRows), numel(vn4));
    for i = 1:numel(perfRows)
        p = perfRows(i);
        rows4(i, :) = {p.metric, p.value, p.ci, p.n, p.level, p.subgroup, p.page, p.section, p.quote, yn(p.review), p.note};
    end
    if isempty(perfRows)
        rows4 = {'', '', '', '', 'Not disclosed', '', '', '', '', 'Yes', ...
            'No numeric performance values were parsed from the extracted text (see D3-D5 in Disclosure Audit). Tables/figures stored as images require OCR or manual reading.'};
    end
    T4 = makeTable(rows4, vn4);

    % ---- 5. Disclosure Audit ----
    vn5 = {'Field_ID', 'Group', 'Field', 'Extracted_Value', 'Disclosure_Status', 'Supporting_Evidence', ...
           'PDF_Page', 'Section_Heuristic', 'Needs_Manual_Review', 'Review_Note'};
    rows5 = cell(numel(results), numel(vn5));
    for i = 1:numel(results)
        r = results(i);
        rows5(i, :) = {r.id, r.group, r.field, r.value, r.status, r.evidence, r.page, r.section, yn(r.review), r.note};
    end
    T5 = makeTable(rows5, vn5);

    % ---- 6. Run Log ----
    rows6 = {'Source PDF', runInfo.pdf; 'Extraction backend', runInfo.backend; ...
             'Page numbers reliable', yn(runInfo.pageAware); 'Pages', sprintf('%d', runInfo.nPages); ...
             'Pages with little/no text', mat2str(runInfo.emptyPages); ...
             'Run date', datestr(now, 'yyyy-mm-dd HH:MM'); ...
             'Note', 'Page numbers are PDF page indices (not printed page labels).'};
    for i = 1:numel(runInfo.warnings), rows6(end+1, :) = {sprintf('Warning/check %d', i), runInfo.warnings{i}}; end %#ok<AGROW>
    T6 = makeTable(rows6, {'Item', 'Detail'});

    % ---- write ----
    try
        if exist(xlsxFile, 'file'), delete(xlsxFile); end
        writetable(T1, xlsxFile, 'Sheet', 'Product Summary');
        writetable(T2, xlsxFile, 'Sheet', 'Scanner Configurations');
        writetable(T3, xlsxFile, 'Sheet', 'Validation Dataset');
        writetable(T4, xlsxFile, 'Sheet', 'Performance Results');
        writetable(T5, xlsxFile, 'Sheet', 'Disclosure Audit');
        writetable(T6, xlsxFile, 'Sheet', 'Run Log');
    catch ME
        error('audit510k:writeFail', 'Could not write %s (is it open in Excel?): %s', xlsxFile, ME.message);
    end
end

function T = makeTable(rows, varNames)
% Converts everything to char, clips to Excel's cell limit, builds a table.
    if isempty(rows), rows = cell(0, numel(varNames)); end
    if size(rows, 1) == 1 && size(rows, 2) ~= numel(varNames), rows = reshape(rows, 1, []); end
    for i = 1:numel(rows)
        x = rows{i};
        if isnumeric(x), x = num2str(x); elseif islogical(x), x = iff(x, 'Yes', 'No'); elseif ~ischar(x), x = char(x); end
        if numel(x) > 32000, x = [x(1:32000) ' ...[truncated]']; end
        rows{i} = x;
    end
    T = cell2table(rows, 'VariableNames', varNames);
end

function lp = localParams(win, cfg)
% Reconstruction/acquisition values found inside one scanner-mention window.
    clean = @(c) cellfun(@(m) strtrim(regexprep(m, '\s+', ' ')), c, 'UniformOutput', false);
    lp.algorithm = uniqueStable(clean(regexpi(win, cfg.paramPat.algorithm, 'match')));
    lp.kernel    = uniqueStable(clean(regexpi(win, cfg.paramPat.kernel,    'match')));
    lp.strength  = uniqueStable(clean(regexpi(win, cfg.paramPat.strength,  'match')));
    lp.slice     = uniqueStable(clean(regexpi(win, cfg.paramPat.slice,     'match')));
    lp.acq = {};
    for k = 1:size(cfg.acqPatterns, 1)
        m = clean(regexpi(win, cfg.acqPatterns{k,2}, 'match'));
        if ~isempty(m), lp.acq{end+1} = [cfg.acqPatterns{k,1} ': ' joinList(uniqueStable(m), ', ')]; end %#ok<AGROW>
    end
end

function out = iff(cond, a, b)
    if cond, out = a; else, out = b; end
end