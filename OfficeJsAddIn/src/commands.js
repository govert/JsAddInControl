let isTaskpaneVisible = false;
let isToggleInFlight = false;

Office.onReady(() => {
  Office.actions.associate("toggleTaskpane", toggleTaskpane);

  if (Office.addin && Office.addin.onVisibilityModeChanged) {
    Office.addin.onVisibilityModeChanged(handleVisibilityModeChanged);
  }
});

function handleVisibilityModeChanged(args) {
  const mode = `${args.visibilityMode}`.toLowerCase();
  isTaskpaneVisible = mode.includes("taskpane");
}

function toggleTaskpane(event) {
  event.completed();

  if (isToggleInFlight) {
    return;
  }

  isToggleInFlight = true;

  // The ribbon event must complete before Excel reliably applies the task pane visibility change.
  setTimeout(async () => {
    try {
      if (isTaskpaneVisible) {
        await Office.addin.hide();
        isTaskpaneVisible = false;
      } else {
        await Office.addin.showAsTaskpane();
        isTaskpaneVisible = true;
      }
    } catch (error) {
      console.error("Failed to toggle the task pane.", error);
    } finally {
      isToggleInFlight = false;
    }
  }, 0);
}
