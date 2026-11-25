const insertButton = document.getElementById("sqlookup-insert");
const insertError = document.getElementById("sqlookup-error");

insertButton.addEventListener("click", async () => {
  insertError.hidden = true;
  // Only one script runs at a time, so block new clicks until this one finishes.
  insertButton.disabled = true;
  try {
    await xlwings.runScript("insert_sample_sheet");
  } catch (err) {
    // Drop the exception class, e.g., "XlwingsError: "
    insertError.textContent = err.message.replace(/^[\w.]+Error: /, "");
    insertError.hidden = false;
  } finally {
    insertButton.disabled = false;
  }
});
