document.addEventListener('DOMContentLoaded', () => {

  // Fermeture automatique des alertes après 5 secondes
  setTimeout(() => {
    document.querySelectorAll('.zone-alertes .alert').forEach(a => {
      const inst = bootstrap.Alert.getOrCreateInstance(a);
      inst.close();
    });
  }, 5000);

  // Tooltips Bootstrap
  document.querySelectorAll('[title]').forEach(el => {
    if (!el.closest('a[href]') || el.title.length > 3) {
      try { new bootstrap.Tooltip(el, { trigger: 'hover' }); } catch(e) {}
    }
  });

  // Style des radios
  document.querySelectorAll('.option-radio input').forEach(radio => {
    radio.addEventListener('change', () => {
      document.querySelectorAll(`input[name="${radio.name}"]`).forEach(r => {
        r.closest('.option-radio').classList.toggle('selectionne', r.checked);
      });
    });
  });

  // Champs conditionnels
  document.querySelectorAll('input[data-affiche]').forEach(radio => {
    radio.addEventListener('change', () => {
      const cible = document.getElementById(radio.dataset.affiche);
      if (cible) cible.classList.toggle('d-none', radio.value !== 'Oui');
    });
  });

  // Validation numérique
  document.querySelectorAll('input[type="number"]').forEach(input => {
    input.addEventListener('blur', () => {
      const val = parseFloat(input.value);
      const min = parseFloat(input.min);
      const max = parseFloat(input.max);
      if (!isNaN(val)) {
        if (!isNaN(min) && val < min) input.value = min;
        if (!isNaN(max) && val > max) input.value = max;
      }
    });
  });

});
