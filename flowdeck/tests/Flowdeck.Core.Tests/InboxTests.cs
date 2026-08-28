using Flowdeck.Core.Export;
using Flowdeck.Core.Models;
using Flowdeck.Core.Services;
using Flowdeck.Core.Storage;
using Xunit;

namespace Flowdeck.Core.Tests;

/// <summary>
/// The hand-over from another application: a file appears in a folder, its entries end up in
/// the workspace, and the file is filed away rather than deleted.
/// </summary>
public class InboxImporterTests : IDisposable
{
    private readonly string _root = Path.Combine(
        Path.GetTempPath(),
        "flowdeck-inbox-" + Guid.NewGuid().ToString("N"));

    private readonly WorkspaceRepository _repository;
    private readonly InboxImporter _importer;

    public InboxImporterTests()
    {
        Directory.CreateDirectory(_root);

        _repository = new WorkspaceRepository(new JsonWorkspaceStore(Path.Combine(_root, "workspace.json")));
        _repository.LoadAsync().GetAwaiter().GetResult();

        _importer = new InboxImporter(Path.Combine(_root, "inbox"), _repository.ImportAsync)
        {
            // The backoff is what is being relied on, not how long it takes.
            Delay = _ => Task.CompletedTask,
        };

        _importer.EnsureFolders();
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(_root, recursive: true);
        }
        catch (IOException)
        {
        }
    }

    // ---- helpers -----------------------------------------------------------

    private static TodoItem Todo(string title, string? id = null) => new()
    {
        Id = id ?? Guid.NewGuid().ToString("N"),
        Title = title,
        DueAt = new DateTime(2026, 9, 3, 10, 0, 0),
        HasTime = true,
        Tags = { "리포트" },
    };

    /// <summary>Writes a transfer file into the inbox and returns where it landed.</summary>
    private string Drop(string name, params TodoItem[] todos)
    {
        var path = Path.Combine(_importer.Folder, name);
        File.WriteAllText(path, TransferFile.Write(todos, Array.Empty<CalendarEvent>(), DateTimeOffset.Now));
        return path;
    }

    private string[] Names(string folder) =>
        Directory.GetFiles(folder).Select(Path.GetFileName).OfType<string>().OrderBy(n => n).ToArray();

    // ---- the ordinary path -------------------------------------------------

    [Fact]
    public async Task AFileThatArrivesBecomesEntriesAndIsFiledAway()
    {
        var path = Drop("filebox-1.txt", Todo("[읽기] 상반기 시장분석"));

        var result = await _importer.ProcessAsync(path);

        Assert.NotNull(result);
        Assert.Equal(InboxOutcome.Imported, result!.Outcome);
        Assert.Equal(1, result.Import!.TodosAdded);

        Assert.Equal("[읽기] 상반기 시장분석", Assert.Single(_repository.Todos).Title);
        Assert.False(File.Exists(path));
        Assert.Equal(new[] { "filebox-1.txt" }, Names(_importer.DoneFolder));
    }

    [Fact]
    public async Task TheWallClockTimeSurvivesUntouched()
    {
        // The sending side writes a local wall-clock time with no offset. Ten in the morning
        // has to stay ten in the morning: a time zone conversion here would be silent.
        await _importer.ProcessAsync(Drop("filebox-1.txt", Todo("검토")));

        var todo = Assert.Single(_repository.Todos);
        Assert.Equal(new DateTime(2026, 9, 3, 10, 0, 0), todo.DueAt);
        Assert.True(todo.HasTime);
    }

    [Fact]
    public async Task SeveralEntriesInOneFileAllArrive()
    {
        var path = Drop("filebox-1.txt", Todo("하나"), Todo("둘"), Todo("셋"));

        var result = await _importer.ProcessAsync(path);

        Assert.Equal(3, result!.Import!.TodosAdded);
        Assert.Equal(3, _repository.Todos.Count);
    }

    [Fact]
    public async Task WhatArrivedWhileTheAppWasClosedIsPickedUpAtTheNextStart()
    {
        Drop("filebox-1.txt", Todo("하나"));
        Drop("filebox-2.txt", Todo("둘"));
        Drop("filebox-3.txt", Todo("셋"));

        var results = await _importer.DrainAsync();

        Assert.Equal(3, results.Count);
        Assert.All(results, r => Assert.Equal(InboxOutcome.Imported, r.Outcome));
        Assert.Equal(3, _repository.Todos.Count);
        Assert.Empty(Names(_importer.Folder));
        Assert.Equal(3, Names(_importer.DoneFolder).Length);
    }

    // ---- what must not happen ----------------------------------------------

    [Fact]
    public async Task AHalfWrittenFileIsNotRead()
    {
        // The sending side builds under .tmp and renames into place, which is atomic. Until
        // that rename the file is not ours to touch, whatever is in it.
        var path = Path.Combine(_importer.Folder, "filebox-1.tmp");
        File.WriteAllText(path, TransferFile.Write(new[] { Todo("아직") }, Array.Empty<CalendarEvent>(), DateTimeOffset.Now));

        Assert.Null(await _importer.ProcessAsync(path));
        Assert.Empty(await _importer.DrainAsync());

        Assert.Empty(_repository.Todos);
        Assert.True(File.Exists(path));
    }

    [Fact]
    public async Task TheSameEntriesArrivingTwiceStayOneEntry()
    {
        var todo = Todo("한 번만", id: "3f9c1a52b7d4462e8a01c6d5e9f27b41");

        await _importer.ProcessAsync(Drop("filebox-1.txt", todo));
        var second = await _importer.ProcessAsync(Drop("filebox-2.txt", todo));

        Assert.Equal(0, second!.Import!.Added);
        Assert.Equal(1, second.Import.Skipped);
        Assert.Single(_repository.Todos);
    }

    [Fact]
    public async Task AnEntryEditedHereIsNotOverwrittenBySendingItAgain()
    {
        var todo = Todo("원래 제목", id: "3f9c1a52b7d4462e8a01c6d5e9f27b41");
        await _importer.ProcessAsync(Drop("filebox-1.txt", todo));

        var mine = Assert.Single(_repository.Todos);
        mine.Title = "내가 고친 제목";
        await _repository.ToggleTodoAsync(mine.Id, new DateTime(2026, 9, 1, 9, 0, 0));

        await _importer.ProcessAsync(Drop("filebox-2.txt", todo));

        var after = Assert.Single(_repository.Todos);
        Assert.Equal("내가 고친 제목", after.Title);
        Assert.True(after.IsDone);
    }

    [Fact]
    public async Task WhateverWasLinkedOverThereIsNotLinkedHere()
    {
        var todo = Todo("보내온 것");
        todo.ExternalLink = new ExternalLink { EntryId = "저쪽-ID", StoreId = "저쪽-저장소" };

        await _importer.ProcessAsync(Drop("filebox-1.txt", todo));

        Assert.Null(Assert.Single(_repository.Todos).ExternalLink);
    }

    // ---- files that are not ours -------------------------------------------

    [Fact]
    public async Task BrokenJsonIsSetAsideAndTheAppCarriesOn()
    {
        var path = Path.Combine(_importer.Folder, "filebox-1.txt");
        File.WriteAllText(path, "--- 여기서부터는 앱이 읽는 부분입니다. 지우지 마세요 ---\n{ 이건 JSON이");

        var result = await _importer.ProcessAsync(path);

        Assert.Equal(InboxOutcome.Rejected, result!.Outcome);
        Assert.NotNull(result.Problem);
        Assert.Empty(_repository.Todos);
        Assert.Equal(new[] { "filebox-1.txt" }, Names(_importer.FailedFolder));
    }

    [Fact]
    public async Task SomeOtherApplicationsJsonIsSetAside()
    {
        var path = Path.Combine(_importer.Folder, "something-else.txt");
        File.WriteAllText(path, """{ "Format": "something.else", "Version": 1 }""");

        Assert.Equal(InboxOutcome.Rejected, (await _importer.ProcessAsync(path))!.Outcome);
        Assert.Equal(new[] { "something-else.txt" }, Names(_importer.FailedFolder));
    }

    [Fact]
    public async Task AFileFromANewerVersionIsSetAside()
    {
        var path = Path.Combine(_importer.Folder, "filebox-1.txt");
        File.WriteAllText(path, """{ "Format": "flowdeck.transfer", "Version": 99 }""");

        Assert.Equal(InboxOutcome.Rejected, (await _importer.ProcessAsync(path))!.Outcome);
        Assert.Equal(new[] { "filebox-1.txt" }, Names(_importer.FailedFolder));
    }

    [Fact]
    public async Task SomethingTooBigToBeAHandOverIsSetAsideUnread()
    {
        var path = Path.Combine(_importer.Folder, "huge.txt");
        File.WriteAllText(path, new string('x', (int)InboxImporter.SizeCeiling + 1));

        Assert.Equal(InboxOutcome.Rejected, (await _importer.ProcessAsync(path))!.Outcome);
        Assert.Equal(new[] { "huge.txt" }, Names(_importer.FailedFolder));
    }

    [Fact]
    public async Task AFileStillBeingWrittenIsLeftForNextTime()
    {
        var path = Drop("filebox-1.txt", Todo("잠김"));

        using (File.Open(path, FileMode.Open, FileAccess.Read, FileShare.None))
        {
            var result = await _importer.ProcessAsync(path);

            Assert.Equal(InboxOutcome.Deferred, result!.Outcome);
            Assert.True(File.Exists(path));
            Assert.Empty(_repository.Todos);
        }

        // Released: the next scan takes it.
        Assert.Equal(InboxOutcome.Imported, (await _importer.ProcessAsync(path))!.Outcome);
        Assert.Single(_repository.Todos);
    }

    // ---- filing away -------------------------------------------------------

    [Fact]
    public async Task TwoHandOversWithTheSameNameBothSurvive()
    {
        await _importer.ProcessAsync(Drop("filebox.txt", Todo("첫 번째")));
        await _importer.ProcessAsync(Drop("filebox.txt", Todo("두 번째")));
        await _importer.ProcessAsync(Drop("filebox.txt", Todo("세 번째")));

        Assert.Equal(
            new[] { "filebox (2).txt", "filebox (3).txt", "filebox.txt" },
            Names(_importer.DoneFolder));
        Assert.Equal(3, _repository.Todos.Count);
    }

    [Fact]
    public async Task WhatWasFiledAwayIsNotReadAgain()
    {
        await _importer.ProcessAsync(Drop("filebox-1.txt", Todo("한 번")));

        // 처리됨 and 실패 sit inside the inbox: whoever scans it must not walk into them.
        Assert.Empty(_importer.Pending());
        Assert.Empty(await _importer.DrainAsync());
        Assert.Single(_repository.Todos);
    }

    /// <summary>
    /// The file the sending side was told to write, character for character, as it appears in
    /// the bridge specification. It is the whole of the agreement between the two applications
    /// and neither can see the other's code, so it is pinned here: renaming a property on
    /// <see cref="TodoItem"/> would otherwise break the hand-over silently, since a key that
    /// does not match is not an error to the deserialiser — the value simply never arrives.
    /// </summary>
    [Fact]
    public async Task TheFileTheSendingSideWasToldToWriteIsUnderstood()
    {
        var path = Path.Combine(_importer.Folder, "filebox-20260827-143210-3f9c1a52.txt");
        File.WriteAllText(path, """
            FileBox → Flowdeck · 2026-08-27 14:32 · 1건

            - [할일] 2026 상반기 시장분석 리포트 · 9월 3일 10:00 #리포트 #읽기

            --- 여기서부터는 앱이 읽는 부분입니다. 지우지 마세요 ---
            {
              "Format": "flowdeck.transfer",
              "Version": 1,
              "ExportedAt": "2026-08-27T14:32:10+09:00",
              "Todos": [
                {
                  "Id": "3f9c1a52b7d4462e8a01c6d5e9f27b41",
                  "Title": "[읽기] 2026 상반기 시장분석 리포트",
                  "Notes": "파일: C:\\Users\\andrew\\Documents\\관리함\\2026_상반기_시장분석.pdf\n출처: FileBox 특별규칙 \"리포트 읽기\"",
                  "DueAt": "2026-09-03T10:00:00",
                  "HasTime": true,
                  "Priority": "Normal",
                  "Tags": ["리포트", "읽기"],
                  "IsDone": false,
                  "CreatedAt": "2026-08-27T14:32:10",
                  "UpdatedAt": "2026-08-27T14:32:10",
                  "SourceText": "2026_상반기_시장분석.pdf",
                  "ReminderMinutesBefore": 30
                }
              ],
              "Events": []
            }
            """);

        Assert.Equal(InboxOutcome.Imported, (await _importer.ProcessAsync(path))!.Outcome);

        var todo = Assert.Single(_repository.Todos);
        Assert.Equal("3f9c1a52b7d4462e8a01c6d5e9f27b41", todo.Id);
        Assert.Equal("[읽기] 2026 상반기 시장분석 리포트", todo.Title);
        Assert.Contains(@"파일: C:\Users\andrew\Documents\관리함\2026_상반기_시장분석.pdf", todo.Notes);
        Assert.Equal(new DateTime(2026, 9, 3, 10, 0, 0), todo.DueAt);
        Assert.Equal(DateTimeKind.Unspecified, todo.DueAt!.Value.Kind);
        Assert.True(todo.HasTime);
        Assert.Equal(Priority.Normal, todo.Priority);
        Assert.Equal(new[] { "리포트", "읽기" }, todo.Tags);
        Assert.False(todo.IsDone);
        Assert.Equal("2026_상반기_시장분석.pdf", todo.SourceText);
        Assert.Equal(30, todo.ReminderMinutesBefore);

        // And the path in that note survives well enough to open the file from the entry.
        Assert.Equal(
            @"C:\Users\andrew\Documents\관리함\2026_상반기_시장분석.pdf",
            AttachedFile.PathIn(todo.Notes));

        // Left out of the file on purpose, and each has to land on its own default rather
        // than on null: the todo goes straight into a workspace that expects them set.
        Assert.False(todo.Recurrence.IsRepeating);
        Assert.Null(todo.ExternalLink);
        Assert.Null(todo.LinkedEventId);
    }

    [Fact]
    public async Task ManyAtOnceAllLandInTheWorkspace()
    {
        var paths = Enumerable.Range(1, 10)
            .Select(n => Drop($"filebox-{n}.txt", Todo($"항목 {n}")))
            .ToList();

        await Task.WhenAll(paths.Select(p => _importer.ProcessAsync(p)));

        Assert.Equal(10, _repository.Todos.Count);

        // And on disk, not merely in memory: two imports overlapping would have each saved
        // the workspace from its own copy, and the later save would have dropped the earlier.
        var reloaded = new WorkspaceRepository(new JsonWorkspaceStore(Path.Combine(_root, "workspace.json")));
        await reloaded.LoadAsync();
        Assert.Equal(10, reloaded.Todos.Count);
    }
}
