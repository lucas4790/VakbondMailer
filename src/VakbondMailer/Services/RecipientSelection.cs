using System.Data;
using System.Linq;
using VakbondMailer.Models;

namespace VakbondMailer.Services;

/// <summary>
/// De ingelezen ledenlijst zoals die in beeld staat, plus wie er aangevinkt is om te mailen.
///
/// De tabel staat 1-op-1 in dezelfde volgorde als de ingelezen lijst — daarom staat sorteren in
/// het scherm uit: rij-index is ontvanger-index. Die koppeling zit hier, want daar is eerder
/// een fout in geslopen (na sorteren werd de verkeerde persoon getoond en getest).
/// </summary>
public sealed class RecipientSelection
{
    public const string SelectionColumnName = "Versturen";

    private RecipientSelection(DataTable table, IReadOnlyList<Recipient> all)
    {
        Table = table;
        All = all;
    }

    public DataTable Table { get; }

    public IReadOnlyList<Recipient> All { get; }

    public static RecipientSelection From(ImportedRecipients imported)
    {
        var table = new DataTable();
        table.Columns.Add(SelectionColumnName, typeof(bool));
        foreach (var header in imported.Headers)
            table.Columns.Add(header);

        foreach (var recipient in imported.Recipients)
        {
            var row = table.NewRow();
            row[SelectionColumnName] = true;
            foreach (var header in imported.Headers)
                row[header] = recipient.Fields.TryGetValue(header, out var value) ? value : string.Empty;
            table.Rows.Add(row);
        }

        return new RecipientSelection(table, imported.Recipients);
    }

    public IReadOnlyList<Recipient> Selected =>
        All.Where((_, index) => index < Table.Rows.Count && Table.Rows[index][SelectionColumnName] is true)
            .ToList();

    /// <summary>De ontvanger op een rij uit de tabel, voor het live voorbeeld.</summary>
    public Recipient? At(int index) =>
        index >= 0 && index < All.Count ? All[index] : All.FirstOrDefault();

    public void SetAll(bool selected)
    {
        foreach (DataRow row in Table.Rows)
            row[SelectionColumnName] = selected;
    }

    /// <summary>Vinkt precies deze ontvangers aan en de rest uit (bv. na een mislukte ronde).</summary>
    public void SelectOnly(IReadOnlyCollection<Recipient> recipients)
    {
        for (var index = 0; index < Table.Rows.Count && index < All.Count; index++)
            Table.Rows[index][SelectionColumnName] = recipients.Contains(All[index]);
    }

    public string CountLabel
    {
        get
        {
            var total = All.Count;
            var selected = Selected.Count;
            return selected == total
                ? $"{total} ontvanger(s) geladen"
                : $"{selected} van {total} geselecteerd";
        }
    }
}
