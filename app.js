const STORAGE_KEY = "course-map-builder-v1";
const APP_VERSION = 1;

const assetConfigs = {
    assessments: {
        title: "Assessments",
        copy: "Ways students show what they have learned in this module.",
        rowLabel: "Assessment",
        typeLabel: "Type",
        typePlaceholder: "Paper, project, quiz, presentation",
        descriptionLabel: "Description",
        descriptionPlaceholder: "Describe the assessment and add a link if useful.",
        linkLabel: "Link",
        linkPlaceholder: "Rubric, prompt, or shared file link"
    },
    materials: {
        title: "Instructional materials",
        copy: "Readings, media, handouts, and other supports students will use in this module.",
        rowLabel: "Material",
        typeLabel: "Type",
        typePlaceholder: "Video, article, PowerPoint, website",
        descriptionLabel: "Title or reference information",
        descriptionPlaceholder: "Name the material or summarize its role in the module.",
        linkLabel: "Website or link",
        linkPlaceholder: "Website, library, or OneDrive link"
    },
    activities: {
        title: "Learning activities",
        copy: "Low-stakes practice that helps students work toward the module objectives, such as a short reflection, discussion, or problem set.",
        rowLabel: "Activity",
        typeLabel: "Type",
        typePlaceholder: "Discussion, reading, reflection, case study",
        descriptionLabel: "Description",
        descriptionPlaceholder: "Describe the learning activity and what learners will do.",
        linkLabel: "Website or link",
        linkPlaceholder: "Website, instructions, or shared file link"
    }
};

const moduleStudioSteps = [
    { id: "summary", navLabel: "Summary", title: "Module summary" },
    { id: "objectives", navLabel: "Objectives", title: "Objectives" },
    { id: "assessments", navLabel: "Assessments", title: "Assessments" },
    { id: "materials", navLabel: "Materials", title: "Materials" },
    { id: "activities", navLabel: "Activities", title: "Activities" }
];

const elements = {
    boundFields: Array.from(document.querySelectorAll("[data-bind]")),
    courseObjectives: document.getElementById("course-objectives"),
    heroMetrics: document.getElementById("hero-metrics"),
    importInput: document.getElementById("import-input"),
    moduleZone: document.getElementById("module-zone"),
    progressFill: document.getElementById("progress-fill"),
    readinessList: document.getElementById("readiness-list"),
    readinessScore: document.getElementById("readiness-score"),
    readinessStatus: document.getElementById("readiness-status"),
    saveStatus: document.getElementById("save-status"),
    studioList: document.getElementById("studio-list")
};

let state = hydrateState(loadState());
const viewState = {
    activeModuleId: null,
    activeModuleStep: moduleStudioSteps[0].id,
    sectionEditors: {}
};

function initialize() {
    bindGlobalEvents();
    renderAll();
    updateSaveStatus(state.meta.savedAt);
}

function bindGlobalEvents() {
    document.addEventListener("click", handleClick);
    document.addEventListener("input", handleInput);
    document.addEventListener("change", handleChange);
    document.addEventListener("keydown", handleKeydown);
    elements.importInput?.addEventListener("change", handleImportFile);
}

function handleClick(event) {
    const actionTrigger = event.target.closest("[data-action]");
    if (!actionTrigger) {
        return;
    }

    const { action, moduleId, objectiveId, assetId, assetKind, studioId, stepId } = actionTrigger.dataset;

    if (action === "open-module-studio") {
        const module = findModule(moduleId);
        if (!module) {
            return;
        }
        clearSectionEditors();
        viewState.activeModuleId = moduleId;
        viewState.activeModuleStep = getSuggestedModuleStep(module);
        renderModules();
        focusActiveModuleStep();
        return;
    }

    if (action === "close-module-studio") {
        const activeModuleId = moduleId || viewState.activeModuleId;
        viewState.activeModuleId = null;
        viewState.activeModuleStep = moduleStudioSteps[0].id;
        clearSectionEditors();
        renderModules();
        if (activeModuleId) {
            focusSelector(`[data-action="open-module-studio"][data-module-id="${activeModuleId}"]`);
        }
        return;
    }

    if (action === "set-module-step") {
        viewState.activeModuleStep = getModuleStudioStep(stepId).id;
        renderModules();
        focusActiveModuleStep();
        return;
    }

    if (action === "module-step-previous" || action === "module-step-next") {
        const stepOffset = action === "module-step-next" ? 1 : -1;
        const nextStepId = getAdjacentModuleStepId(viewState.activeModuleStep, stepOffset);
        if (!nextStepId) {
            return;
        }
        viewState.activeModuleStep = nextStepId;
        renderModules();
        focusActiveModuleStep();
        return;
    }

    if (action === "add-course-objective") {
        const existingEmptyObjective = state.courseObjectives.find((item) => !hasText(item.text));
        if (existingEmptyObjective) {
            focusSelector(`[data-course-objective-id="${existingEmptyObjective.id}"][data-field="text"]`);
            return;
        }
        state.courseObjectives.push(createCourseObjective());
        renderCourseObjectives();
        renderModules();
        refreshDashboard();
        persistState();
        focusSelector(`[data-course-objective-id="${state.courseObjectives.at(-1).id}"][data-field="text"]`);
        return;
    }

    if (action === "remove-course-objective") {
        if (state.courseObjectives.length === 1) {
            return;
        }
        const objective = state.courseObjectives.find((item) => item.id === objectiveId);
        if (objective && hasText(objective.text) && !window.confirm("Remove this course objective?")) {
            return;
        }
        state.courseObjectives = state.courseObjectives.filter((item) => item.id !== objectiveId);
        state.modules.forEach((module) => {
            module.objectives.forEach((moduleObjective) => {
                moduleObjective.alignsTo = moduleObjective.alignsTo.filter((id) => id !== objectiveId);
            });
        });
        renderCourseObjectives();
        renderModules();
        refreshDashboard();
        persistState();
        return;
    }

    if (action === "add-module") {
        const module = createModule();
        state.modules.push(module);
        clearSectionEditors();
        viewState.activeModuleId = module.id;
        viewState.activeModuleStep = moduleStudioSteps[0].id;
        renderModules();
        renderStudioItems();
        refreshDashboard();
        persistState();
        focusSelector(`[data-module-id="${module.id}"][data-module-field="title"]`);
        return;
    }

    if (action === "remove-module") {
        const module = findModule(moduleId);
        if (!module) {
            return;
        }
        if (!window.confirm("Remove this module and everything inside it?")) {
            return;
        }
        state.modules = state.modules.filter((item) => item.id !== moduleId);
        if (viewState.activeModuleId === moduleId) {
            viewState.activeModuleId = null;
        }
        viewState.activeModuleStep = moduleStudioSteps[0].id;
        clearSectionEditorsForModule(moduleId);
        state.studioItems.forEach((item) => {
            if (item.moduleId === moduleId) {
                item.moduleId = "";
            }
        });
        renderModules();
        renderStudioItems();
        refreshDashboard();
        persistState();
        if (!state.modules.length) {
            focusSelector('[data-action="add-module"]');
        }
        return;
    }

    if (action === "add-module-objective") {
        const module = findModule(moduleId);
        if (!module) {
            return;
        }
        const existingEmptyObjective = module.objectives.find((item) => !hasText(item.text));
        if (existingEmptyObjective) {
            viewState.sectionEditors[getSectionEditorKey(moduleId, "objectives")] = existingEmptyObjective.id;
            renderModules();
            focusSelector(`[data-module-id="${moduleId}"][data-module-objective-id="${existingEmptyObjective.id}"][data-field="text"]`);
            return;
        }
        const objective = createModuleObjective();
        module.objectives.push(objective);
        viewState.sectionEditors[getSectionEditorKey(moduleId, "objectives")] = objective.id;
        renderModules();
        refreshDashboard();
        persistState();
        focusSelector(`[data-module-id="${moduleId}"][data-module-objective-id="${objective.id}"][data-field="text"]`);
        return;
    }

    if (action === "edit-module-objective") {
        viewState.sectionEditors[getSectionEditorKey(moduleId, "objectives")] = objectiveId;
        renderModules();
        focusSelector(`[data-module-id="${moduleId}"][data-module-objective-id="${objectiveId}"][data-field="text"]`);
        return;
    }

    if (action === "remove-module-objective") {
        const module = findModule(moduleId);
        if (!module || module.objectives.length === 1) {
            return;
        }
        const objective = findModuleObjective(moduleId, objectiveId);
        if (objective && hasText(objective.text) && !window.confirm("Remove this module objective?")) {
            return;
        }
        module.objectives = module.objectives.filter((item) => item.id !== objectiveId);
        if (viewState.sectionEditors[getSectionEditorKey(moduleId, "objectives")] === objectiveId) {
            viewState.sectionEditors[getSectionEditorKey(moduleId, "objectives")] = "";
        }
        for (const assetKindName of Object.keys(assetConfigs)) {
            module[assetKindName].forEach((asset) => {
                asset.alignsTo = asset.alignsTo.filter((id) => id !== objectiveId);
            });
        }
        renderModules();
        refreshDashboard();
        persistState();
        return;
    }

    if (action === "add-asset-row") {
        const module = findModule(moduleId);
        if (!module) {
            return;
        }
        const existingEmptyAsset = module[assetKind].find((item) => !isFilledAsset(item));
        if (existingEmptyAsset) {
            viewState.sectionEditors[getSectionEditorKey(moduleId, assetKind)] = existingEmptyAsset.id;
            renderModules();
            focusSelector(`[data-module-id="${moduleId}"][data-asset-kind="${assetKind}"][data-asset-id="${existingEmptyAsset.id}"][data-field="type"]`);
            return;
        }
        const asset = createAsset();
        module[assetKind].push(asset);
        viewState.sectionEditors[getSectionEditorKey(moduleId, assetKind)] = asset.id;
        renderModules();
        refreshDashboard();
        persistState();
        focusSelector(`[data-module-id="${moduleId}"][data-asset-kind="${assetKind}"][data-asset-id="${asset.id}"][data-field="type"]`);
        return;
    }

    if (action === "edit-asset-row") {
        viewState.sectionEditors[getSectionEditorKey(moduleId, assetKind)] = assetId;
        renderModules();
        focusSelector(`[data-module-id="${moduleId}"][data-asset-kind="${assetKind}"][data-asset-id="${assetId}"][data-field="type"]`);
        return;
    }

    if (action === "remove-asset-row") {
        const module = findModule(moduleId);
        if (!module) {
            return;
        }
        const asset = findAsset(moduleId, assetKind, assetId);
        if (asset && isFilledAsset(asset) && !window.confirm(`Remove this ${assetConfigs[assetKind].rowLabel.toLowerCase()} row?`)) {
            return;
        }
        module[assetKind] = module[assetKind].filter((item) => item.id !== assetId);
        if (viewState.sectionEditors[getSectionEditorKey(moduleId, assetKind)] === assetId) {
            viewState.sectionEditors[getSectionEditorKey(moduleId, assetKind)] = "";
        }
        renderModules();
        refreshDashboard();
        persistState();
        return;
    }

    if (action === "close-section-editor") {
        viewState.sectionEditors[getSectionEditorKey(moduleId, assetKind)] = "";
        renderModules();
        return;
    }

    if (action === "add-studio-item") {
        const item = createStudioItem();
        state.studioItems.push(item);
        renderStudioItems();
        refreshDashboard();
        persistState();
        focusSelector(`[data-studio-id="${item.id}"][data-field="title"]`);
        return;
    }

    if (action === "remove-studio-item") {
        if (state.studioItems.length === 1) {
            return;
        }
        const item = findStudioItem(studioId);
        if (item && isFilledStudioItem(item) && !window.confirm("Remove this studio tracking row?")) {
            return;
        }
        state.studioItems = state.studioItems.filter((entry) => entry.id !== studioId);
        renderStudioItems();
        refreshDashboard();
        persistState();
        return;
    }

    if (action === "download-word") {
        downloadWordExport();
        return;
    }

    if (action === "print-pdf") {
        openPrintPreview();
        return;
    }

    if (action === "download-json") {
        downloadJsonBackup();
        return;
    }

    if (action === "import-json") {
        elements.importInput.click();
        return;
    }

    if (action === "clear-draft") {
        if (!window.confirm("Clear the saved draft in this browser and start fresh?")) {
            return;
        }
        localStorage.removeItem(STORAGE_KEY);
        state = createDefaultState();
        resetViewState();
        renderAll();
        updateSaveStatus(null);
        focusSelector('[data-bind="course.name"]');
        return;
    }
}

function handleKeydown(event) {
    if (event.key !== "Escape" || !viewState.activeModuleId) {
        return;
    }
    const activeModuleId = viewState.activeModuleId;
    viewState.activeModuleId = null;
    viewState.activeModuleStep = moduleStudioSteps[0].id;
    clearSectionEditors();
    renderModules();
    focusSelector(`[data-action="open-module-studio"][data-module-id="${activeModuleId}"]`);
}

function handleInput(event) {
    const target = event.target;

    if (target.matches("[data-bind]")) {
        setByPath(state, target.dataset.bind, target.value);
        refreshDashboard();
        persistState();
        return;
    }

    if (target.dataset.courseObjectiveId) {
        const objective = state.courseObjectives.find((item) => item.id === target.dataset.courseObjectiveId);
        if (!objective) {
            return;
        }
        objective[target.dataset.field] = target.value;
        renderModules();
        refreshDashboard();
        persistState();
        return;
    }

    if (target.dataset.moduleId && target.dataset.moduleField) {
        const module = findModule(target.dataset.moduleId);
        if (!module) {
            return;
        }
        module[target.dataset.moduleField] = target.value;
        refreshModuleCard(target.dataset.moduleId);
        refreshDashboard();
        persistState();
        return;
    }

    if (target.dataset.moduleObjectiveId && target.type !== "checkbox") {
        const objective = findModuleObjective(target.dataset.moduleId, target.dataset.moduleObjectiveId);
        if (!objective) {
            return;
        }
        objective[target.dataset.field] = target.value;
        refreshModuleCard(target.dataset.moduleId);
        refreshDashboard();
        persistState();
        return;
    }

    if (target.dataset.assetId && target.type !== "checkbox") {
        const asset = findAsset(target.dataset.moduleId, target.dataset.assetKind, target.dataset.assetId);
        if (!asset) {
            return;
        }
        asset[target.dataset.field] = target.value;
        refreshModuleCard(target.dataset.moduleId);
        refreshDashboard();
        persistState();
        return;
    }

    if (target.dataset.studioId) {
        const item = findStudioItem(target.dataset.studioId);
        if (!item) {
            return;
        }
        item[target.dataset.field] = target.value;
        refreshDashboard();
        persistState();
    }
}

function handleChange(event) {
    const target = event.target;

    if (target.dataset.alignTarget === "course-objective") {
        const objective = findModuleObjective(target.dataset.moduleId, target.dataset.moduleObjectiveId);
        if (!objective) {
            return;
        }
        objective.alignsTo = toggleInArray(objective.alignsTo, target.value, target.checked);
        refreshModuleCard(target.dataset.moduleId);
        refreshDashboard();
        persistState();
        return;
    }

    if (target.dataset.alignTarget === "module-objective") {
        const asset = findAsset(target.dataset.moduleId, target.dataset.assetKind, target.dataset.assetId);
        if (!asset) {
            return;
        }
        asset.alignsTo = toggleInArray(asset.alignsTo, target.value, target.checked);
        refreshModuleCard(target.dataset.moduleId);
        refreshDashboard();
        persistState();
    }
}

function handleImportFile(event) {
    const [file] = event.target.files || [];
    if (!file) {
        return;
    }

    const reader = new FileReader();
    reader.onload = () => {
        try {
            const imported = JSON.parse(String(reader.result));
            state = hydrateState(imported);
            renderAll();
            persistState();
        } catch (error) {
            window.alert("That file could not be read as a course map backup.");
        } finally {
            elements.importInput.value = "";
        }
    };
    reader.readAsText(file);
}

function renderAll() {
    renderBoundFields();
    renderCourseObjectives();
    renderModules();
    renderStudioItems();
    refreshDashboard();
}

function renderBoundFields() {
    for (const field of elements.boundFields) {
        field.value = getByPath(state, field.dataset.bind) || "";
    }
}

function renderCourseObjectives() {
    elements.courseObjectives.innerHTML = state.courseObjectives.map((objective, index) => `
        <article class="item-card">
            <div class="item-head">
                <div>
                    <p class="item-label">CO${index + 1}</p>
                </div>
                <button class="btn btn-small btn-ghost" type="button" data-action="remove-course-objective" data-objective-id="${objective.id}" ${state.courseObjectives.length === 1 ? "disabled" : ""}>Remove</button>
            </div>
            <label class="field">
                <span>Objective</span>
                <textarea class="objective-text" rows="2" placeholder="Describe a measurable course outcome." data-course-objective-id="${objective.id}" data-field="text">${escapeHtml(objective.text)}</textarea>
            </label>
        </article>
    `).join("");
}

function renderModules() {
    if (!viewState.activeModuleId) {
        elements.moduleZone.innerHTML = renderModuleList();
        syncModuleStudioMode();
        return;
    }

    const module = findModule(viewState.activeModuleId);
    if (!module) {
        viewState.activeModuleId = null;
        viewState.activeModuleStep = moduleStudioSteps[0].id;
        clearSectionEditors();
        elements.moduleZone.innerHTML = renderModuleList();
        syncModuleStudioMode();
        return;
    }

    viewState.activeModuleStep = getModuleStudioStep(viewState.activeModuleStep).id;
    elements.moduleZone.innerHTML = renderModuleStudio(module, state.modules.findIndex((entry) => entry.id === module.id));
    syncModuleStudioMode();
}

function renderModuleList() {
    return `
        <div class="module-list">
            ${state.modules.length
                ? state.modules.map((module, index) => renderModuleListCard(module, index)).join("")
                : `
                    <div class="empty-state">
                        <p>No modules added yet. Add a module to begin building the course sequence.</p>
                    </div>
                `}
        </div>
        ${renderAddRow({
            label: "Add module",
            action: "add-module"
        })}
    `;
}

function renderModuleListCard(module, index) {
    const moduleName = hasText(module.title) ? module.title : `Untitled module ${index + 1}`;
    const summary = hasText(module.overview) ? truncate(module.overview, 120) : "No module overview yet.";
    const objectiveCount = module.objectives.filter((objective) => hasText(objective.text)).length;
    const plannedItems = Object.keys(assetConfigs).reduce((sum, kind) => sum + module[kind].filter(isFilledAsset).length, 0);
    const metaMarkup = objectiveCount || plannedItems
        ? `
            <div class="module-list-card__meta">
                <span class="summary-pill">${objectiveCount} objective${objectiveCount === 1 ? "" : "s"}</span>
                <span class="summary-pill">${plannedItems} planned item${plannedItems === 1 ? "" : "s"}</span>
            </div>
        `
        : `<p class="module-list-card__status">Not started yet.</p>`;

    return `
        <article class="module-list-card">
            <div class="module-list-card__copy">
                <p class="summary-kicker">Module ${index + 1}</p>
                <h3 data-module-list-heading="${module.id}">${escapeHtml(moduleName)}</h3>
                <p class="module-list-card__summary">${escapeHtml(summary)}</p>
            </div>
            ${metaMarkup}
            <div class="module-list-card__actions">
                <button class="btn btn-secondary" type="button" data-action="open-module-studio" data-module-id="${module.id}">Open module</button>
                <button class="btn btn-small btn-ghost" type="button" data-action="remove-module" data-module-id="${module.id}">Remove</button>
            </div>
        </article>
    `;
}

function renderModuleStudio(module, index) {
    const moduleName = hasText(module.title) ? module.title : `Untitled module ${index + 1}`;
    const activeStep = getModuleStudioStep(viewState.activeModuleStep);
    const stepIndex = moduleStudioSteps.findIndex((step) => step.id === activeStep.id);
    const previousStepId = getAdjacentModuleStepId(activeStep.id, -1);
    const nextStepId = getAdjacentModuleStepId(activeStep.id, 1);

    return `
        <div class="module-studio-shell">
            <button class="module-studio__backdrop" type="button" data-action="close-module-studio" data-module-id="${module.id}" aria-label="Close module workspace"></button>
            <section class="module-studio" role="dialog" aria-modal="true" aria-labelledby="module-heading-${module.id}">
                <div class="module-studio__frame">
                    <header class="module-studio__header">
                        <div class="module-studio__bar">
                            <button class="btn btn-secondary btn-small" type="button" data-action="close-module-studio" data-module-id="${module.id}">Return to module list</button>
                            <div class="module-studio__bar-meta">
                                <p class="module-studio__step-note">Step ${stepIndex + 1} of ${moduleStudioSteps.length}</p>
                                <button class="btn btn-small btn-ghost" type="button" data-action="remove-module" data-module-id="${module.id}">Remove module</button>
                            </div>
                        </div>
                        <div class="module-studio__identity">
                            <p class="summary-kicker">Module ${index + 1}</p>
                            <h3 id="module-heading-${module.id}" data-module-heading="${module.id}">${escapeHtml(moduleName)}</h3>
                            <p class="module-studio__save-note">Changes save automatically in this browser.</p>
                        </div>
                    </header>

                    <nav class="module-step-nav" aria-label="Module steps">
                        ${moduleStudioSteps.map((step, navIndex) => `
                            <button
                                class="module-step-btn ${activeStep.id === step.id ? "is-active" : ""}"
                                type="button"
                                data-action="set-module-step"
                                data-module-id="${module.id}"
                                data-step-id="${step.id}"
                            >
                                <span class="module-step-btn__number">${navIndex + 1}</span>
                                <span class="module-step-btn__label">${step.navLabel}</span>
                            </button>
                        `).join("")}
                    </nav>

                    <div class="module-studio__panel">
                        ${renderModuleStudioStep(module, activeStep.id)}
                    </div>

                    <div class="module-studio__footer">
                        <div class="module-studio__footer-nav">
                            ${previousStepId
                                ? `<button class="btn btn-secondary" type="button" data-action="module-step-previous" data-module-id="${module.id}">Previous</button>`
                                : `<span class="module-studio__footer-spacer" aria-hidden="true"></span>`}
                            ${nextStepId
                                ? `<button class="btn btn-primary" type="button" data-action="module-step-next" data-module-id="${module.id}">Next</button>`
                                : `<button class="btn btn-primary" type="button" data-action="close-module-studio" data-module-id="${module.id}">Finish module</button>`}
                        </div>
                    </div>
                </div>
            </section>
        </div>
    `;
}

function renderModuleStudioStep(module, stepId) {
    if (stepId === "summary") {
        return renderModuleOverviewSection(module);
    }
    if (stepId === "objectives") {
        return renderObjectivesStudioSection(module);
    }
    if (stepId === "assessments") {
        return renderAssetStudioSection(module, "assessments", assetConfigs.assessments);
    }
    if (stepId === "materials") {
        return renderAssetStudioSection(module, "materials", assetConfigs.materials);
    }
    return renderAssetStudioSection(module, "activities", assetConfigs.activities);
}

function renderModuleOverviewSection(module) {
    return `
        <section class="studio-section">
            <div class="studio-section__head">
                <div>
                    <h4>Module summary</h4>
                    <p class="studio-section__copy">Name the module and briefly describe what it covers and why it matters.</p>
                </div>
            </div>

            <div class="field-grid module-basics">
                <label class="field">
                    <span>Module title</span>
                    <input type="text" placeholder="Human-Nature Relationships" data-module-id="${module.id}" data-module-field="title" value="${escapeHtml(module.title)}">
                </label>
                <label class="field">
                    <span>Module overview</span>
                    <textarea rows="3" placeholder="Summarize what this module covers and why it matters." data-module-id="${module.id}" data-module-field="overview">${escapeHtml(module.overview)}</textarea>
                </label>
            </div>
        </section>
    `;
}

function renderObjectivesStudioSection(module) {
    const activeObjectiveId = getActiveSectionEditorId(module.id, "objectives");
    const activeObjective = module.objectives.find((objective) => objective.id === activeObjectiveId);
    const visibleObjectives = module.objectives.filter((objective) => objective.id !== activeObjectiveId);

    return `
        <section class="studio-section">
            <div class="studio-section__head">
                <div>
                    <h4>Objectives</h4>
                    <p class="studio-section__copy">Write measurable module-level outcomes and align them to the course objectives.</p>
                </div>
            </div>

            <div class="studio-list">
                ${visibleObjectives.map((objective) => renderObjectiveStudioRow(module, objective, module.objectives.findIndex((entry) => entry.id === objective.id), activeObjectiveId)).join("")}
            </div>

            ${activeObjective ? renderObjectiveEditor(module, activeObjective) : ""}

            ${renderAddRow({
                label: "Add module objective",
                action: "add-module-objective",
                moduleId: module.id
            })}
        </section>
    `;
}

function renderObjectiveStudioRow(module, objective, index, activeObjectiveId) {
    const preview = hasText(objective.text) ? truncate(objective.text, 96) : "Objective not yet written";
    return `
        <article class="studio-row ${activeObjectiveId === objective.id ? "is-active" : ""}">
            <div class="studio-row__copy">
                <p class="item-label">MO${index + 1}</p>
                <p class="studio-row__summary">${escapeHtml(preview)}</p>
            </div>
            <div class="studio-row__actions">
                <button class="btn btn-small btn-secondary" type="button" data-action="edit-module-objective" data-module-id="${module.id}" data-objective-id="${objective.id}">Edit</button>
                <button class="btn btn-small btn-ghost" type="button" data-action="remove-module-objective" data-module-id="${module.id}" data-objective-id="${objective.id}" ${module.objectives.length === 1 ? "disabled" : ""}>Remove</button>
            </div>
        </article>
    `;
}

function renderObjectiveEditor(module, objective) {
    const objectiveIndex = module.objectives.findIndex((item) => item.id === objective.id);

    return `
        <section class="studio-editor">
            <div class="studio-editor__head">
                <p class="item-label">MO${objectiveIndex + 1}</p>
                <button class="btn btn-small btn-ghost" type="button" data-action="close-section-editor" data-module-id="${module.id}" data-asset-kind="objectives">Close</button>
            </div>

            <label class="field">
                <span>Objective</span>
                <textarea class="objective-text" rows="2" placeholder="Write a concise, measurable objective." data-module-id="${module.id}" data-module-objective-id="${objective.id}" data-field="text">${escapeHtml(objective.text)}</textarea>
            </label>

            <div class="editor-checklist">
                <span class="field-label">Aligned course objectives</span>
                ${renderCourseObjectiveChoices(module.id, objective)}
            </div>
        </section>
    `;
}

function renderAssetStudioSection(module, assetKind, config) {
    const activeAssetId = getActiveSectionEditorId(module.id, assetKind);
    const activeAsset = module[assetKind].find((asset) => asset.id === activeAssetId);
    const visibleAssets = module[assetKind].filter((asset) => asset.id !== activeAssetId);

    return `
        <section class="studio-section">
            <div class="studio-section__head">
                <div>
                    <h4>${config.title}</h4>
                    <p class="studio-section__copy">${config.copy}</p>
                </div>
            </div>

            <div class="studio-list">
                ${visibleAssets.length
                    ? visibleAssets.map((asset) => renderAssetStudioRow(module, assetKind, config, asset, module[assetKind].findIndex((entry) => entry.id === asset.id), activeAssetId)).join("")
                    : activeAsset
                        ? ""
                        : `
                        <div class="empty-state">
                            <p>No ${config.rowLabel.toLowerCase()}s added yet.</p>
                        </div>
                    `}
            </div>

            ${activeAsset ? renderAssetEditor(module, assetKind, config, activeAsset) : ""}

            ${renderAddRow({
                label: `Add ${config.rowLabel.toLowerCase()}`,
                action: "add-asset-row",
                moduleId: module.id,
                assetKind
            })}
        </section>
    `;
}

function renderAssetStudioRow(module, assetKind, config, asset, index, activeAssetId) {
    const preview = getAssetPreview(config, asset);
    return `
        <article class="studio-row ${activeAssetId === asset.id ? "is-active" : ""}">
            <div class="studio-row__copy">
                <p class="item-label">${config.rowLabel} ${index + 1}</p>
                <p class="studio-row__summary">${escapeHtml(preview)}</p>
            </div>
            <div class="studio-row__actions">
                <button class="btn btn-small btn-secondary" type="button" data-action="edit-asset-row" data-module-id="${module.id}" data-asset-kind="${assetKind}" data-asset-id="${asset.id}">Edit</button>
                <button class="btn btn-small btn-ghost" type="button" data-action="remove-asset-row" data-module-id="${module.id}" data-asset-kind="${assetKind}" data-asset-id="${asset.id}">Remove</button>
            </div>
        </article>
    `;
}

function renderAssetEditor(module, assetKind, config, asset) {
    const assetIndex = module[assetKind].findIndex((item) => item.id === asset.id);

    return `
        <section class="studio-editor">
            <div class="studio-editor__head">
                <p class="item-label">${config.rowLabel} ${assetIndex + 1}</p>
                <button class="btn btn-small btn-ghost" type="button" data-action="close-section-editor" data-module-id="${module.id}" data-asset-kind="${assetKind}">Close</button>
            </div>

            <div class="asset-grid">
                <label class="field">
                    <span>${config.typeLabel}</span>
                    <input type="text" placeholder="${config.typePlaceholder}" data-module-id="${module.id}" data-asset-kind="${assetKind}" data-asset-id="${asset.id}" data-field="type" value="${escapeHtml(asset.type)}">
                </label>
                <label class="field asset-description">
                    <span>${config.descriptionLabel}</span>
                    <textarea rows="3" placeholder="${config.descriptionPlaceholder}" data-module-id="${module.id}" data-asset-kind="${assetKind}" data-asset-id="${asset.id}" data-field="description">${escapeHtml(asset.description)}</textarea>
                </label>
                <label class="field">
                    <span>${config.linkLabel}</span>
                    <input type="text" placeholder="${config.linkPlaceholder}" data-module-id="${module.id}" data-asset-kind="${assetKind}" data-asset-id="${asset.id}" data-field="link" value="${escapeHtml(asset.link)}">
                </label>
            </div>

            <div class="editor-checklist">
                <span class="field-label">Aligned module objectives</span>
                ${renderModuleObjectiveChoices(module, assetKind, asset)}
            </div>
        </section>
    `;
}

function renderStudioItems() {
    if (!elements.studioList) {
        return;
    }
    elements.studioList.innerHTML = state.studioItems.map((item, index) => `
        <article class="item-card">
            <div class="item-head">
                <div>
                    <p class="item-label">Studio ${index + 1}</p>
                    <p class="item-copy">Track only assets that need studio coordination.</p>
                </div>
                <button class="btn btn-small btn-ghost" type="button" data-action="remove-studio-item" data-studio-id="${item.id}" ${state.studioItems.length === 1 ? "disabled" : ""}>Remove</button>
            </div>
            <div class="studio-grid">
                <label class="field">
                    <span>Module</span>
                    <select data-studio-id="${item.id}" data-field="moduleId">
                        <option value="">Select module</option>
                        ${state.modules.map((module, moduleIndex) => `
                            <option value="${module.id}" ${item.moduleId === module.id ? "selected" : ""}>${escapeHtml(getModuleLabel(module, moduleIndex))}</option>
                        `).join("")}
                    </select>
                </label>
                <label class="field">
                    <span>Video title</span>
                    <input type="text" placeholder="Intro lecture: climate justice" data-studio-id="${item.id}" data-field="title" value="${escapeHtml(item.title)}">
                </label>
                <label class="field">
                    <span>Date recorded</span>
                    <input type="date" data-studio-id="${item.id}" data-field="date" value="${escapeHtml(item.date)}">
                </label>
                <label class="field">
                    <span>Freshservice ticket</span>
                    <input type="text" placeholder="FS-10428" data-studio-id="${item.id}" data-field="ticket" value="${escapeHtml(item.ticket)}">
                </label>
                <label class="field">
                    <span>Notes</span>
                    <input type="text" placeholder="Needs lower-third and screen capture." data-studio-id="${item.id}" data-field="notes" value="${escapeHtml(item.notes)}">
                </label>
            </div>
        </article>
    `).join("");
}

function refreshDashboard() {
    if (!elements.readinessScore && !elements.readinessStatus && !elements.readinessList && !elements.heroMetrics && !elements.progressFill) {
        return;
    }
    const readiness = getReadiness();
    if (elements.readinessScore) {
        elements.readinessScore.textContent = `${readiness.score}%`;
    }
    if (elements.readinessStatus) {
        elements.readinessStatus.textContent = readiness.status;
    }
    if (elements.progressFill) {
        elements.progressFill.style.width = `${readiness.score}%`;
    }
    if (elements.readinessList) {
        elements.readinessList.innerHTML = readiness.items.map((item) => `
            <li>${escapeHtml(item.message)}</li>
        `).join("");
    }
    if (elements.heroMetrics) {
        elements.heroMetrics.innerHTML = buildMetrics().map((metric) => `
            <article class="metric-card">
                <span>${escapeHtml(metric.label)}</span>
                <strong>${escapeHtml(String(metric.value))}</strong>
            </article>
        `).join("");
    }
}

function refreshModuleCard(moduleId) {
    const module = findModule(moduleId);
    if (!module) {
        return;
    }
    const headingNode = document.querySelector(`[data-module-heading="${moduleId}"]`);
    const listHeadingNode = document.querySelector(`[data-module-list-heading="${moduleId}"]`);
    const headingText = hasText(module.title) ? module.title : `Untitled module ${state.modules.findIndex((entry) => entry.id === moduleId) + 1}`;
    if (headingNode) {
        headingNode.textContent = headingText;
    }
    if (listHeadingNode) {
        listHeadingNode.textContent = headingText;
    }
}

function syncModuleStudioMode() {
    document.body.classList.toggle("is-studio-open", Boolean(viewState.activeModuleId));
}

function renderAddRow({ label, action, moduleId = "", assetKind = "" }) {
    const moduleAttribute = moduleId ? ` data-module-id="${moduleId}"` : "";
    const assetAttribute = assetKind ? ` data-asset-kind="${assetKind}"` : "";

    return `
        <div class="add-row">
            <button class="btn btn-add add-row__button" type="button" data-action="${action}"${moduleAttribute}${assetAttribute}>
                <span class="add-row__icon" aria-hidden="true">+</span>
                <span>${escapeHtml(label)}</span>
            </button>
        </div>
    `;
}

function getSectionEditorKey(moduleId, section) {
    return `${moduleId}:${section}`;
}

function getActiveSectionEditorId(moduleId, section) {
    return viewState.sectionEditors[getSectionEditorKey(moduleId, section)] || "";
}

function clearSectionEditorsForModule(moduleId) {
    Object.keys(viewState.sectionEditors).forEach((key) => {
        if (key.startsWith(`${moduleId}:`)) {
            delete viewState.sectionEditors[key];
        }
    });
}

function clearSectionEditors() {
    viewState.sectionEditors = {};
}

function resetViewState() {
    viewState.activeModuleId = null;
    viewState.activeModuleStep = moduleStudioSteps[0].id;
    viewState.sectionEditors = {};
}

function getModuleStudioStep(stepId) {
    return moduleStudioSteps.find((step) => step.id === stepId) || moduleStudioSteps[0];
}

function getAdjacentModuleStepId(stepId, offset) {
    const currentIndex = moduleStudioSteps.findIndex((step) => step.id === stepId);
    const nextStep = moduleStudioSteps[currentIndex + offset];
    return nextStep ? nextStep.id : "";
}

function getSuggestedModuleStep(module) {
    if (!hasText(module.title) && !hasText(module.overview)) {
        return "summary";
    }
    if (!module.objectives.some((objective) => hasText(objective.text))) {
        return "objectives";
    }
    if (!module.assessments.some(isFilledAsset)) {
        return "assessments";
    }
    if (!module.materials.some(isFilledAsset)) {
        return "materials";
    }
    if (!module.activities.some(isFilledAsset)) {
        return "activities";
    }
    return "summary";
}

function focusActiveModuleStep() {
    const moduleId = viewState.activeModuleId;
    if (!moduleId) {
        return;
    }
    const stepSelectors = {
        summary: `[data-module-id="${moduleId}"][data-module-field="title"]`,
        objectives: `[data-action="add-module-objective"][data-module-id="${moduleId}"]`,
        assessments: `[data-action="add-asset-row"][data-module-id="${moduleId}"][data-asset-kind="assessments"]`,
        materials: `[data-action="add-asset-row"][data-module-id="${moduleId}"][data-asset-kind="materials"]`,
        activities: `[data-action="add-asset-row"][data-module-id="${moduleId}"][data-asset-kind="activities"]`
    };
    focusSelector(stepSelectors[viewState.activeModuleStep] || `[data-action="close-module-studio"][data-module-id="${moduleId}"]`);
}

function renderCourseObjectiveChoices(moduleId, objective) {
    const alignableObjectives = getAlignableCourseObjectives();
    if (!alignableObjectives.length) {
        return `<p class="empty-hint">Draft at least one course objective first.</p>`;
    }

    return `
        <div class="alignment-list">
            ${alignableObjectives.map(({ objective: courseObjective, index }) => {
                const checked = objective.alignsTo.includes(courseObjective.id);
                return `
                    <label class="alignment-option">
                        <input type="checkbox" value="${courseObjective.id}" data-align-target="course-objective" data-module-id="${moduleId}" data-module-objective-id="${objective.id}" ${checked ? "checked" : ""}>
                        <span>${escapeHtml(getCourseObjectiveLabel(courseObjective, index))}</span>
                    </label>
                `;
            }).join("")}
        </div>
    `;
}

function renderModuleObjectiveChoices(module, assetKind, asset) {
    const alignableObjectives = getAlignableModuleObjectives(module);
    if (!alignableObjectives.length) {
        return `<p class="empty-hint">Draft at least one module objective first.</p>`;
    }

    return `
        <div class="alignment-list">
            ${alignableObjectives.map(({ objective, index }) => {
                const checked = asset.alignsTo.includes(objective.id);
                return `
                    <label class="alignment-option">
                        <input type="checkbox" value="${objective.id}" data-align-target="module-objective" data-module-id="${module.id}" data-asset-kind="${assetKind}" data-asset-id="${asset.id}" ${checked ? "checked" : ""}>
                        <span>${escapeHtml(getModuleObjectiveLabel(objective, index))}</span>
                    </label>
                `;
            }).join("")}
        </div>
    `;
}

function buildMetrics() {
    const filledCourseObjectives = state.courseObjectives.filter((objective) => hasText(objective.text)).length;
    const alignmentLinks = state.modules.reduce((sum, module) => {
        const moduleObjectiveLinks = module.objectives.reduce((inner, objective) => inner + objective.alignsTo.length, 0);
        const assetLinks = Object.keys(assetConfigs).reduce((inner, kind) => inner + module[kind].reduce((assetSum, asset) => assetSum + asset.alignsTo.length, 0), 0);
        return sum + moduleObjectiveLinks + assetLinks;
    }, 0);
    const studioItems = state.studioItems.filter(isFilledStudioItem).length;

    return [
        { label: "Course objectives drafted", value: filledCourseObjectives },
        { label: "Modules in plan", value: state.modules.length },
        { label: "Alignment links", value: alignmentLinks },
        { label: "Studio items", value: studioItems }
    ];
}

function getReadiness() {
    let passed = 0;
    let total = 0;
    const issues = [];

    const check = (condition, message) => {
        total += 1;
        if (condition) {
            passed += 1;
        } else {
            issues.push({ level: "warning", message });
        }
    };

    check(hasText(state.course.name) && hasText(state.course.code) && hasText(state.course.instructor), "Fill in the course name, code, and instructor.");
    check(hasText(state.course.overview), "Add a course overview so the map has a clear through-line.");
    check(state.courseObjectives.some((objective) => hasText(objective.text)), "Add at least one course objective.");

    state.modules.forEach((module, index) => {
        const label = getModuleLabel(module, index);
        check(hasText(module.title), `Give ${label} a title.`);
        check(module.objectives.some((objective) => hasText(objective.text)), `Add at least one module objective in ${label}.`);
        check(module.objectives.filter((objective) => hasText(objective.text)).every((objective) => objective.alignsTo.length > 0), `Align each drafted objective in ${label} to a course objective.`);
        check(module.assessments.some(isFilledAsset), `Add at least one assessment to ${label}.`);
        check(module.materials.some(isFilledAsset), `Add at least one instructional material to ${label}.`);
        check(module.activities.some(isFilledAsset), `Add at least one learning activity to ${label}.`);
        check(getFilledAssets(module).every((asset) => asset.alignsTo.length > 0), `Connect each drafted row in ${label} to one or more module objectives.`);
    });

    const score = total === 0 ? 0 : Math.round((passed / total) * 100);
    const status = score >= 92
        ? "The map has the core pieces in place and is ready for export."
        : score >= 70
            ? "The draft is in good shape. Focus on the next few gaps only."
            : "Start with the first few missing pieces and keep moving top to bottom.";

    return {
        score,
        status,
        items: issues.length ? issues.slice(0, 3) : [{ level: "good", message: "Everything essential is in place for a clean draft export." }]
    };
}

function getModuleHealth(module) {
    const checks = [
        hasText(module.title),
        hasText(module.overview),
        module.objectives.some((objective) => hasText(objective.text)),
        module.assessments.some(isFilledAsset),
        module.materials.some(isFilledAsset),
        module.activities.some(isFilledAsset)
    ];
    const passed = checks.filter(Boolean).length;
    const nextItems = [];

    if (!hasText(module.title)) {
        nextItems.push("add a title");
    }
    if (!hasText(module.overview)) {
        nextItems.push("add a short overview");
    }
    if (!module.objectives.some((objective) => hasText(objective.text))) {
        nextItems.push("draft at least one objective");
    }
    if (!module.assessments.some(isFilledAsset)) {
        nextItems.push("add an assessment");
    }
    if (!module.materials.some(isFilledAsset)) {
        nextItems.push("add a material");
    }
    if (!module.activities.some(isFilledAsset)) {
        nextItems.push("add an activity");
    }

    return {
        summary: `${passed}/${checks.length} checkpoints`,
        detail: nextItems.length ? `Next: ${nextItems.slice(0, 2).join(" and ")}.` : "This module has the core planning pieces in place."
    };
}

function downloadWordExport() {
    const html = buildExportHtml({ forPrint: false });
    const blob = new Blob(["\ufeff", html], { type: "application/msword;charset=utf-8" });
    downloadBlob(blob, `${buildFileStem()}.doc`);
}

function openPrintPreview() {
    const preview = window.open("", "_blank", "noopener,noreferrer");
    if (!preview) {
        window.alert("Please allow pop-ups so the print preview can open.");
        return;
    }
    preview.onload = () => {
        preview.focus();
        preview.print();
    };
    preview.document.open();
    preview.document.write(buildExportHtml({ forPrint: true }));
    preview.document.close();
}

function downloadJsonBackup() {
    const blob = new Blob([JSON.stringify({ ...state, version: APP_VERSION }, null, 2)], { type: "application/json;charset=utf-8" });
    downloadBlob(blob, `${buildFileStem()}.json`);
}

// Build a clean, document-style export from the structured state instead of printing the raw form UI.
function buildExportHtml({ forPrint }) {
    const moduleSections = state.modules.map((module, index) => renderExportModule(module, index)).join("");

    return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>${escapeHtml(buildDocumentTitle())}</title>
    <style>
        body {
            margin: 0;
            padding: 32px;
            font-family: "Segoe UI", Arial, sans-serif;
            color: #1f2d27;
            background: #fffdfa;
        }
        .doc-shell {
            max-width: 980px;
            margin: 0 auto;
        }
        .doc-header {
            padding: 22px 26px;
            border-radius: 18px;
            background: linear-gradient(135deg, #214433, #537c63);
            color: #fffdfa;
        }
        .doc-header h1 {
            margin: 0;
            font-size: 30px;
        }
        .doc-header p {
            margin: 8px 0 0;
            color: rgba(255, 252, 244, 0.88);
        }
        .doc-section {
            margin-top: 28px;
        }
        .doc-section h2 {
            margin: 0 0 12px;
            padding-bottom: 10px;
            border-bottom: 2px solid #d2c17f;
            color: #214433;
            font-size: 24px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 14px;
        }
        th, td {
            border: 1px solid #d9ddd8;
            padding: 10px 12px;
            vertical-align: top;
            text-align: left;
            font-size: 14px;
            line-height: 1.45;
        }
        th {
            background: #f1f5f0;
            color: #214433;
        }
        .meta-grid {
            display: grid;
            gap: 12px;
            grid-template-columns: repeat(2, minmax(0, 1fr));
        }
        .meta-card {
            padding: 14px 16px;
            border: 1px solid #d9ddd8;
            border-radius: 14px;
            background: #fff;
        }
        .meta-card strong {
            display: block;
            margin-bottom: 6px;
            color: #214433;
            font-size: 13px;
            letter-spacing: 0.08em;
            text-transform: uppercase;
        }
        .body-copy {
            margin: 0;
            padding: 16px 18px;
            border-radius: 14px;
            border: 1px solid #d9ddd8;
            background: #fff;
            line-height: 1.65;
        }
        .module-block {
            margin-top: 22px;
            padding: 20px;
            border: 1px solid #d9ddd8;
            border-radius: 18px;
            background: #fff;
        }
        .module-block h3 {
            margin: 0 0 8px;
            color: #214433;
            font-size: 22px;
        }
        .small-note {
            margin: 0;
            color: #567065;
            font-size: 13px;
        }
        a {
            color: #1f6a4d;
            text-decoration: none;
        }
        ${forPrint ? `
        @page {
            margin: 0.55in;
        }
        ` : ""}
    </style>
</head>
<body>
    <div class="doc-shell">
        <header class="doc-header">
            <h1>${escapeHtml(buildDocumentTitle())}</h1>
            <p>Generated from the Course Map Builder on ${escapeHtml(new Date().toLocaleDateString())}</p>
        </header>

        <section class="doc-section">
            <h2>Course Setup</h2>
            <div class="meta-grid">
                ${renderMetaCard("Course name", state.course.name)}
                ${renderMetaCard("Course code", state.course.code)}
                ${renderMetaCard("Instructor", state.course.instructor)}
                ${renderMetaCard("Development term", state.course.term)}
                ${renderMetaCard("Instructional designer", joinNonEmpty([state.course.designer, state.course.designerEmail]))}
                ${renderMetaCard("Canvas URL", linkify(state.course.canvasUrl), true)}
                ${renderMetaCard("Other tools", state.course.tools)}
                ${renderMetaCard("Textbook citation", state.course.textbook)}
            </div>
        </section>

        <section class="doc-section">
            <h2>Foundations</h2>
            <p class="body-copy">${nl2br(state.course.overview || "No course overview added.")}</p>
            <table>
                <thead>
                    <tr>
                        <th style="width: 100px;">CO #</th>
                        <th>Course Objective</th>
                    </tr>
                </thead>
                <tbody>
                    ${renderCourseObjectiveTable()}
                </tbody>
            </table>
        </section>

        <section class="doc-section">
            <h2>Modules</h2>
            ${moduleSections}
        </section>
    </div>
</body>
</html>`;
}

function renderExportModule(module, index) {
    const nonEmptyObjectives = module.objectives.filter((objective) => hasText(objective.text));
    const objectiveRows = nonEmptyObjectives.map((objective, objectiveIndex) => `
        <tr>
            <td>${escapeHtml(getModuleObjectiveCode(objectiveIndex))}</td>
            <td>${escapeHtml(objective.text)}</td>
            <td>${escapeHtml(renderCourseObjectiveReferences(objective.alignsTo))}</td>
        </tr>
    `).join("") || `<tr><td colspan="3">No module objectives added.</td></tr>`;

    const assessmentRows = renderExportAssetRows(module, "assessments");
    const materialRows = renderExportAssetRows(module, "materials");
    const activityRows = renderExportAssetRows(module, "activities");

    return `
        <article class="module-block">
            <h3>${escapeHtml(getModuleLabel(module, index))}</h3>
            <p class="small-note">${escapeHtml(module.overview || "No module overview added.")}</p>

            <table>
                <thead>
                    <tr>
                        <th style="width: 120px;">MO #</th>
                        <th>Module Objective</th>
                        <th style="width: 180px;">Aligned Course Objective(s)</th>
                    </tr>
                </thead>
                <tbody>
                    ${objectiveRows}
                </tbody>
            </table>

            <table>
                <thead>
                    <tr>
                        <th style="width: 160px;">Assessment Type</th>
                        <th>Description</th>
                        <th style="width: 200px;">Aligned Module Objective(s)</th>
                    </tr>
                </thead>
                <tbody>
                    ${assessmentRows}
                </tbody>
            </table>

            <table>
                <thead>
                    <tr>
                        <th style="width: 160px;">Material Type</th>
                        <th>Title / Reference</th>
                        <th style="width: 200px;">Website / Link</th>
                        <th style="width: 200px;">Aligned Module Objective(s)</th>
                    </tr>
                </thead>
                <tbody>
                    ${materialRows}
                </tbody>
            </table>

            <table>
                <thead>
                    <tr>
                        <th style="width: 160px;">Activity Type</th>
                        <th>Description</th>
                        <th style="width: 200px;">Website / Link</th>
                        <th style="width: 200px;">Aligned Module Objective(s)</th>
                    </tr>
                </thead>
                <tbody>
                    ${activityRows}
                </tbody>
            </table>
        </article>
    `;
}

function renderExportAssetRows(module, assetKind) {
    const filledAssets = module[assetKind].filter(isFilledAsset);
    if (!filledAssets.length) {
        const colspan = assetKind === "assessments" ? 3 : 4;
        return `<tr><td colspan="${colspan}">No ${assetConfigs[assetKind].rowLabel.toLowerCase()} rows added.</td></tr>`;
    }

    return filledAssets.map((asset) => {
        const cells = [
            `<td>${escapeHtml(asset.type || "—")}</td>`,
            `<td>${nl2br(asset.description || "—")}</td>`
        ];

        if (assetKind !== "assessments") {
            cells.push(`<td>${linkify(asset.link)}</td>`);
        }

        cells.push(`<td>${escapeHtml(renderModuleObjectiveReferences(module, asset.alignsTo))}</td>`);

        return `<tr>${cells.join("")}</tr>`;
    }).join("");
}

function renderCourseObjectiveTable() {
    const objectives = state.courseObjectives.filter((objective) => hasText(objective.text));
    if (!objectives.length) {
        return `<tr><td colspan="2">No course objectives added.</td></tr>`;
    }
    return objectives.map((objective, index) => `
        <tr>
            <td>${escapeHtml(getCourseObjectiveCode(index))}</td>
            <td>${escapeHtml(objective.text)}</td>
        </tr>
    `).join("");
}

function renderMetaCard(label, value, allowHtml = false) {
    const renderedValue = allowHtml ? value : escapeHtml(value || "—");
    return `
        <div class="meta-card">
            <strong>${escapeHtml(label)}</strong>
            <div>${renderedValue}</div>
        </div>
    `;
}

function renderCourseObjectiveReferences(ids) {
    const labels = ids.map((id) => {
        const objectiveIndex = state.courseObjectives.findIndex((item) => item.id === id);
        if (objectiveIndex === -1) {
            return null;
        }
        return getCourseObjectiveCode(objectiveIndex);
    }).filter(Boolean);
    return labels.length ? labels.join(", ") : "—";
}

function renderModuleObjectiveReferences(module, ids) {
    const labels = ids.map((id) => {
        const objectiveIndex = module.objectives.findIndex((item) => item.id === id);
        if (objectiveIndex === -1) {
            return null;
        }
        return getModuleObjectiveCode(objectiveIndex);
    }).filter(Boolean);
    return labels.length ? labels.join(", ") : "—";
}

function persistState() {
    state.meta.savedAt = new Date().toISOString();
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ ...state, version: APP_VERSION }));
    updateSaveStatus(state.meta.savedAt);
}

function loadState() {
    try {
        const raw = localStorage.getItem(STORAGE_KEY);
        return raw ? JSON.parse(raw) : null;
    } catch (error) {
        return null;
    }
}

function createDefaultState() {
    return {
        version: APP_VERSION,
        course: {
            name: "",
            code: "",
            instructor: "",
            term: "",
            designer: "",
            designerEmail: "",
            canvasUrl: "",
            tools: "",
            textbook: "",
            overview: ""
        },
        courseObjectives: [createCourseObjective()],
        modules: [],
        studioItems: [],
        meta: {
            savedAt: null
        }
    };
}

function hydrateState(input) {
    const base = createDefaultState();
    if (!input || typeof input !== "object") {
        return base;
    }

    const hydrated = {
        version: APP_VERSION,
        course: {
            ...base.course,
            ...(input.course || {})
        },
        courseObjectives: Array.isArray(input.courseObjectives) && input.courseObjectives.length
            ? input.courseObjectives.map((objective) => createCourseObjective(objective))
            : base.courseObjectives,
        modules: Array.isArray(input.modules) && input.modules.length
            ? input.modules.map((module) => createModule(module))
            : base.modules,
        studioItems: Array.isArray(input.studioItems) && input.studioItems.length
            ? input.studioItems.map((item) => createStudioItem(item))
            : base.studioItems,
        meta: {
            savedAt: input.meta?.savedAt || null
        }
    };

    hydrated.courseObjectives = normalizeCourseObjectives(hydrated.courseObjectives);
    const validCourseObjectiveIds = new Set(
        hydrated.courseObjectives
            .filter((objective) => hasText(objective.text))
            .map((objective) => objective.id)
    );
    hydrated.modules = hydrated.modules.map((module) => normalizeModule(module, validCourseObjectiveIds));

    return hydrated;
}

function createCourseObjective(values = {}) {
    return {
        id: values.id || createId("co"),
        text: values.text || ""
    };
}

function createModule(values = {}) {
    return {
        id: values.id || createId("module"),
        title: values.title || "",
        overview: values.overview || "",
        objectives: Array.isArray(values.objectives) && values.objectives.length
            ? values.objectives.map((objective) => createModuleObjective(objective))
            : [createModuleObjective()],
        assessments: Array.isArray(values.assessments) && values.assessments.length
            ? values.assessments.map((asset) => createAsset(asset))
            : [],
        materials: Array.isArray(values.materials) && values.materials.length
            ? values.materials.map((asset) => createAsset(asset))
            : [],
        activities: Array.isArray(values.activities) && values.activities.length
            ? values.activities.map((asset) => createAsset(asset))
            : []
    };
}

function normalizeCourseObjectives(objectives) {
    const filled = objectives.filter((objective) => hasText(objective.text));
    if (!filled.length) {
        return [createCourseObjective(objectives[0] || {})];
    }
    return filled;
}

function normalizeModule(module, validCourseObjectiveIds = new Set()) {
    const normalizedObjectives = module.objectives
        .filter((objective) => hasText(objective.text))
        .map((objective) => ({
            ...objective,
            alignsTo: objective.alignsTo.filter((id) => validCourseObjectiveIds.has(id))
        }));
    const fallbackObjective = createModuleObjective(module.objectives[0] || {});
    fallbackObjective.alignsTo = [];
    const objectives = normalizedObjectives.length ? normalizedObjectives : [fallbackObjective];
    const validModuleObjectiveIds = new Set(normalizedObjectives.map((objective) => objective.id));

    return {
        ...module,
        objectives,
        assessments: module.assessments
            .filter(isFilledAsset)
            .map((asset) => ({
                ...asset,
                alignsTo: asset.alignsTo.filter((id) => validModuleObjectiveIds.has(id))
            })),
        materials: module.materials
            .filter(isFilledAsset)
            .map((asset) => ({
                ...asset,
                alignsTo: asset.alignsTo.filter((id) => validModuleObjectiveIds.has(id))
            })),
        activities: module.activities
            .filter(isFilledAsset)
            .map((asset) => ({
                ...asset,
                alignsTo: asset.alignsTo.filter((id) => validModuleObjectiveIds.has(id))
            }))
    };
}

function createModuleObjective(values = {}) {
    return {
        id: values.id || createId("mo"),
        text: values.text || "",
        alignsTo: Array.isArray(values.alignsTo) ? values.alignsTo.filter(Boolean) : []
    };
}

function createAsset(values = {}) {
    return {
        id: values.id || createId("asset"),
        type: values.type || "",
        description: values.description || "",
        link: values.link || "",
        alignsTo: Array.isArray(values.alignsTo) ? values.alignsTo.filter(Boolean) : []
    };
}

function createStudioItem(values = {}) {
    return {
        id: values.id || createId("studio"),
        moduleId: values.moduleId || "",
        title: values.title || "",
        date: values.date || "",
        ticket: values.ticket || "",
        notes: values.notes || ""
    };
}

function createId(prefix) {
    return `${prefix}-${Math.random().toString(36).slice(2, 10)}-${Date.now().toString(36)}`;
}

function findModule(moduleId) {
    return state.modules.find((module) => module.id === moduleId);
}

function findModuleObjective(moduleId, objectiveId) {
    return findModule(moduleId)?.objectives.find((objective) => objective.id === objectiveId);
}

function findAsset(moduleId, assetKind, assetId) {
    return findModule(moduleId)?.[assetKind].find((asset) => asset.id === assetId);
}

function findStudioItem(studioId) {
    return state.studioItems.find((item) => item.id === studioId);
}

function getFilledAssets(module) {
    return Object.keys(assetConfigs).flatMap((kind) => module[kind].filter(isFilledAsset));
}

function getByPath(object, path) {
    return path.split(".").reduce((value, key) => value?.[key], object);
}

function setByPath(object, path, nextValue) {
    const keys = path.split(".");
    const lastKey = keys.pop();
    let cursor = object;
    for (const key of keys) {
        cursor = cursor[key];
    }
    cursor[lastKey] = nextValue;
}

function toggleInArray(values, value, shouldInclude) {
    const nextValues = new Set(values);
    if (shouldInclude) {
        nextValues.add(value);
    } else {
        nextValues.delete(value);
    }
    return Array.from(nextValues);
}

function buildDocumentTitle() {
    return joinNonEmpty([state.course.code, state.course.name, "Course Map"]) || "Course Map";
}

function buildFileStem() {
    const raw = buildDocumentTitle().toLowerCase();
    return raw.replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "") || "course-map";
}

function getCourseObjectiveCode(index) {
    return `CO${index + 1}`;
}

function getAlignableCourseObjectives() {
    return state.courseObjectives
        .map((objective, index) => ({ objective, index }))
        .filter(({ objective }) => hasText(objective.text));
}

function getCourseObjectiveCodesForIds(ids) {
    return ids.map((id) => {
        const objectiveIndex = state.courseObjectives.findIndex((item) => item.id === id && hasText(item.text));
        if (objectiveIndex === -1) {
            return null;
        }
        return getCourseObjectiveCode(objectiveIndex);
    }).filter(Boolean);
}

function getCourseObjectiveLabel(objective, index) {
    const preview = hasText(objective.text) ? `: ${truncate(objective.text, 40)}` : "";
    return `${getCourseObjectiveCode(index)}${preview}`;
}

function getModuleObjectiveCode(index) {
    return `MO${index + 1}`;
}

function getAlignableModuleObjectives(module) {
    return module.objectives
        .map((objective, index) => ({ objective, index }))
        .filter(({ objective }) => hasText(objective.text));
}

function getModuleObjectiveCodesForIds(module, ids) {
    return ids.map((id) => {
        const objectiveIndex = module.objectives.findIndex((item) => item.id === id && hasText(item.text));
        if (objectiveIndex === -1) {
            return null;
        }
        return getModuleObjectiveCode(objectiveIndex);
    }).filter(Boolean);
}

function getModuleObjectiveLabel(objective, index) {
    const preview = hasText(objective.text) ? `: ${truncate(objective.text, 36)}` : "";
    return `${getModuleObjectiveCode(index)}${preview}`;
}

function getAlignmentSummary(labels) {
    if (!labels.length) {
        return "None selected";
    }
    if (labels.length <= 2) {
        return labels.join(", ");
    }
    return `${labels.slice(0, 2).join(", ")} +${labels.length - 2} more`;
}

function getAssetPreview(config, asset) {
    const previewSource = joinNonEmpty([asset.type, asset.description, asset.link]);
    if (!hasText(previewSource)) {
        return `${config.rowLabel} not yet added`;
    }
    return truncate(previewSource, 84);
}

function getModuleLabel(module, index) {
    return hasText(module.title) ? `Module ${index + 1}: ${module.title}` : `Module ${index + 1}`;
}

function getStudioModuleLabel(moduleId) {
    const index = state.modules.findIndex((module) => module.id === moduleId);
    if (index === -1) {
        return "—";
    }
    return getModuleLabel(state.modules[index], index);
}

function formatDate(value) {
    if (!value) {
        return "—";
    }
    const date = new Date(`${value}T00:00:00`);
    if (Number.isNaN(date.getTime())) {
        return value;
    }
    return date.toLocaleDateString();
}

function updateSaveStatus(timestamp) {
    if (!elements.saveStatus) {
        return;
    }
    elements.saveStatus.textContent = timestamp ? `Saved ${new Date(timestamp).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })}` : "Not saved yet";
}

function isFilledAsset(asset) {
    return hasText(asset.type) || hasText(asset.description) || hasText(asset.link);
}

function isFilledStudioItem(item) {
    return hasText(item.title) || hasText(item.ticket) || hasText(item.notes) || hasText(item.date);
}

function hasText(value) {
    return Boolean(String(value || "").trim());
}

function joinNonEmpty(values) {
    return values.filter(hasText).join(" • ");
}

function truncate(value, limit) {
    const clean = String(value || "").trim();
    if (clean.length <= limit) {
        return clean;
    }
    return `${clean.slice(0, limit - 1).trim()}...`;
}

function escapeHtml(value) {
    return String(value || "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
}

function linkify(value) {
    if (!hasText(value)) {
        return "—";
    }
    const trimmed = value.trim();
    if (/^https?:\/\//i.test(trimmed) || /^mailto:/i.test(trimmed)) {
        const safe = escapeHtml(trimmed);
        return `<a href="${safe}" target="_blank" rel="noopener noreferrer">${safe}</a>`;
    }
    return nl2br(trimmed);
}

function nl2br(value) {
    return escapeHtml(value || "—").replace(/\n/g, "<br>");
}

function downloadBlob(blob, filename) {
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
}

function focusSelector(selector) {
    const element = document.querySelector(selector);
    if (element) {
        element.focus();
        element.select?.();
    }
}

initialize();
