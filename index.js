import { initUi, commandChooseCl, commandChooseWs, commandRunComparison, commandGenerateRevisedWs } from './app.js';
Office.onReady(() => {
    if (document.readyState === 'loading')
        document.addEventListener('DOMContentLoaded', initUi);
    else
        initUi();
    Office.actions.associate('commandChooseCl', commandChooseCl);
    Office.actions.associate('commandChooseWs', commandChooseWs);
    Office.actions.associate('commandRunComparison', commandRunComparison);
    Office.actions.associate('commandGenerateRevisedWs', commandGenerateRevisedWs);
});
