// Den Excel-Helfer als schmales Fenster am rechten Bildschirmrand öffnen, neben Excel.
document.querySelectorAll("a[data-popup]").forEach((link) => {
  link.addEventListener("click", (event) => {
    const width = 430;
    const features = `popup,width=${width},height=${screen.availHeight},left=${Math.max(0, screen.availWidth - width)},top=0`;
    const opened = window.open(link.href, "voiceagent-helfer", features);
    if (opened) {
      event.preventDefault();
      opened.focus();
    }
  });
});
