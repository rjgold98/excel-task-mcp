namespace ExcelTask.Core;

/// <summary>
/// The one place that knows which payload belongs to which operation kind.
///
/// The union used to be restated as a hand-kept parallel array in the engine, and an operation
/// missing from it was not merely uncounted - it became unreachable, because a request carrying
/// only that payload failed the arity check and never reached its own validation. Two operations
/// shipped that way. The guard added afterwards worked by maintaining a second hand-kept list in a
/// test, which protects the one site and does not extend to the others.
///
/// A switch expression with no default arm is what makes this different: adding a member to
/// <see cref="ExcelOperationKind"/> without adding an arm here is a compile error, not a silent
/// omission discovered by a user. The compiler is the checklist.
/// </summary>
public static class OperationCatalog
{
    /// <summary>Every kind the server defines, in declaration order.</summary>
    public static IReadOnlyList<ExcelOperationKind> AllKinds { get; } = Enum.GetValues<ExcelOperationKind>();

    /// <summary>The payload this kind requires, or null when the request did not supply it.</summary>
    public static object? PayloadFor(ExcelOperation operation, ExcelOperationKind kind)
    {
        ArgumentNullException.ThrowIfNull(operation);

        // No default arm, deliberately - that is what makes a forgotten member a build error
        // (CS8509). Only CS8524 is suppressed, which is the separate complaint that a value outside
        // the enum's named members is unhandled; requests are rejected by Enum.IsDefined before they
        // reach here, and a `_ =>` arm added to silence it would also swallow the omission this
        // type exists to catch.
#pragma warning disable CS8524
        return kind switch
        {
            ExcelOperationKind.CopyExhibit => operation.CopyExhibit,
            ExcelOperationKind.RepairExistingWorksheet => operation.RepairExistingWorksheet,
            ExcelOperationKind.ExtendFormulaSeries => operation.ExtendFormulaSeries,
            ExcelOperationKind.EditMacroProcedure => operation.EditMacroProcedure,
            ExcelOperationKind.AuditWorkbookFlows => operation.AuditWorkbookFlows,
            ExcelOperationKind.ReadWorksheetRange => operation.ReadWorksheetRange,
            ExcelOperationKind.WriteWorksheetValues => operation.WriteWorksheetValues,
            ExcelOperationKind.WriteWorksheetFormulas => operation.WriteWorksheetFormulas,
            ExcelOperationKind.FindReplace => operation.FindReplace,
            ExcelOperationKind.Create => operation.Create,
            ExcelOperationKind.SetRangeFormat => operation.SetRangeFormat,
            ExcelOperationKind.ScanWorkbookStructure => operation.ScanWorkbookStructure,
            ExcelOperationKind.ManageTable => operation.ManageTable,
            ExcelOperationKind.ManageQuery => operation.ManageQuery,
            ExcelOperationKind.ManageModelMeasure => operation.ManageModelMeasure,
            ExcelOperationKind.ManageModelRelationship => operation.ManageModelRelationship
        };
#pragma warning restore CS8524
    }

    /// <summary>
    /// How many payloads the request supplied. Exactly one is the contract; this counts by asking
    /// each kind for its own payload rather than by listing them again.
    /// </summary>
    public static int SuppliedPayloadCount(ExcelOperation operation)
    {
        ArgumentNullException.ThrowIfNull(operation);

        var count = 0;
        foreach (var kind in AllKinds)
        {
            if (PayloadFor(operation, kind) is not null) count++;
        }

        return count;
    }
}
