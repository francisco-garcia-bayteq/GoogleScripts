const Calculation = (function () {
  function applyWeightedAverage(record, weights, threshold) {
    const grade1 = record.grade1 || 0;
    const grade2 = record.grade2 || 0;
    const grade3 = record.grade3 || 0;
    const weight1 = weights.grade1 != null ? weights.grade1 : 0;
    const weight2 = weights.grade2 != null ? weights.grade2 : 0;
    const weight3 = weights.grade3 != null ? weights.grade3 : 0;
    const finalGrade = Number((grade1 * weight1 + grade2 * weight2 + grade3 * weight3).toFixed(2));
    const status = finalGrade >= threshold ? 'Approved' : 'Requires Improvement';

    return Object.assign({}, record, {
      finalGrade,
      status,
      approvalThreshold: threshold,
    });
  }

  return {
    applyWeightedAverage,
  };
})();
