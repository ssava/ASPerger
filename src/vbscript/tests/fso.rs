use super::*;
    // ===== FILESYSTEMOBJECT + TEXTSTREAM =====


    #[test]
    fn test_fso_createobject() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "Set fso = CreateObject(\"Scripting.FileSystemObject\")",
                &mut ctx,
            )
            .unwrap();
        let fso = ctx.get_variable("fso");
        assert!(fso.is_some());
        assert!(matches!(fso.unwrap(), VBValue::Object(_)));
    }

    #[test]
    fn test_fso_fileexists() {
        let path = tmp_path("fileexists.txt");
        cleanup_path(&path);
        assert!(!std::path::Path::new(&path).exists());

        // Create the file
        std::fs::File::create(&path).unwrap();
        assert!(std::path::Path::new(&path).exists());

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\nx = fso.FileExists(\"{}\")",
                    path
                ),
                &mut ctx,
            )
            .unwrap();
        assert_eq!(ctx.get_variable("x"), Some(&VBValue::Boolean(true)));

        cleanup_path(&path);
    }

    #[test]
    fn test_fso_folderexists() {
        let path = tmp_path("folderexists");
        cleanup_path(&path);
        std::fs::create_dir_all(&path).unwrap();

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\nx = fso.FolderExists(\"{}\")",
                    path
                ),
                &mut ctx,
            )
            .unwrap();
        assert_eq!(ctx.get_variable("x"), Some(&VBValue::Boolean(true)));

        // Non-existent folder
        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\nx = fso.FolderExists(\"{}_nonexistent\")",
                    path
                ),
                &mut ctx,
            )
            .unwrap();
        assert_eq!(ctx.get_variable("x"), Some(&VBValue::Boolean(false)));

        cleanup_path(&path);
    }

    #[test]
    fn test_fso_createtextfile_and_readall() {
        let path = tmp_path("create_read.txt");
        cleanup_path(&path);

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        // Create, write, close
        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     Set ts = fso.CreateTextFile(\"{}\", True)\n\
                     ts.Write \"Hello, World!\"\n\
                     ts.Close",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        // Open and read all
        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     Set ts = fso.OpenTextFile(\"{}\", 1)\n\
                     content = ts.ReadAll()\n\
                     ts.Close",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        assert_eq!(
            ctx.get_variable("content"),
            Some(&VBValue::String("Hello, World!".to_string()))
        );

        cleanup_path(&path);
    }

    #[test]
    fn test_fso_writeline_and_readline() {
        let path = tmp_path("writeline.txt");
        cleanup_path(&path);

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        // Write multiple lines
        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     Set ts = fso.CreateTextFile(\"{}\", True)\n\
                     ts.WriteLine \"Line 1\"\n\
                     ts.WriteLine \"Line 2\"\n\
                     ts.WriteLine \"Line 3\"\n\
                     ts.Close",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        // Read them back
        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     Set ts = fso.OpenTextFile(\"{}\", 1)\n\
                     line1 = ts.ReadLine()\n\
                     line2 = ts.ReadLine()\n\
                     line3 = ts.ReadLine()\n\
                     ts.Close",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        assert_eq!(
            ctx.get_variable("line1"),
            Some(&VBValue::String("Line 1".to_string()))
        );
        assert_eq!(
            ctx.get_variable("line2"),
            Some(&VBValue::String("Line 2".to_string()))
        );
        assert_eq!(
            ctx.get_variable("line3"),
            Some(&VBValue::String("Line 3".to_string()))
        );

        cleanup_path(&path);
    }

    #[test]
    fn test_fso_read_n_characters() {
        let path = tmp_path("readchars.txt");
        cleanup_path(&path);

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        std::fs::write(&path, "ABCDEFGHIJ").unwrap();

        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     Set ts = fso.OpenTextFile(\"{}\", 1)\n\
                     part = ts.Read(4)\n\
                     ts.Close",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        assert_eq!(
            ctx.get_variable("part"),
            Some(&VBValue::String("ABCD".to_string()))
        );

        cleanup_path(&path);
    }

    #[test]
    fn test_fso_atendofstream() {
        let path = tmp_path("atend.txt");
        cleanup_path(&path);

        std::fs::write(&path, "Hello").unwrap();

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     Set ts = fso.OpenTextFile(\"{}\", 1)\n\
                     ' Should not be at end initially\n\
                     initial = ts.AtEndOfStream\n\
                     content = ts.ReadAll()\n\
                     ' Should be at end after reading all\n\
                     after = ts.AtEndOfStream\n\
                     ts.Close",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        assert_eq!(ctx.get_variable("initial"), Some(&VBValue::Boolean(false)));
        assert_eq!(ctx.get_variable("after"), Some(&VBValue::Boolean(true)));

        cleanup_path(&path);
    }

    #[test]
    fn test_fso_path_functions() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        interp
            .execute(
                "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                 parent = fso.GetParentFolderName(\"/a/b/c.txt\")\n\
                 fname = fso.GetFileName(\"/a/b/c.txt\")\n\
                 ext = fso.GetExtensionName(\"/a/b/c.txt\")\n\
                 base = fso.GetBaseName(\"/a/b/c.txt\")\n\
                 built = fso.BuildPath(\"/a/b\", \"c.txt\")",
                &mut ctx,
            )
            .unwrap();

        assert_eq!(
            ctx.get_variable("parent"),
            Some(&VBValue::String("/a/b".to_string()))
        );
        assert_eq!(
            ctx.get_variable("fname"),
            Some(&VBValue::String("c.txt".to_string()))
        );
        assert_eq!(
            ctx.get_variable("ext"),
            Some(&VBValue::String("txt".to_string()))
        );
        assert_eq!(
            ctx.get_variable("base"),
            Some(&VBValue::String("c".to_string()))
        );
        assert_eq!(
            ctx.get_variable("built"),
            Some(&VBValue::String("/a/b/c.txt".to_string()))
        );
    }

    #[test]
    fn test_fso_copyfile() {
        let src = tmp_path("copy_src.txt");
        let dst = tmp_path("copy_dst.txt");
        cleanup_path(&src);
        cleanup_path(&dst);

        std::fs::write(&src, "Copy test content").unwrap();

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     fso.CopyFile \"{}\", \"{}\", True",
                    src, dst
                ),
                &mut ctx,
            )
            .unwrap();

        assert!(std::path::Path::new(&dst).exists());
        let content = std::fs::read_to_string(&dst).unwrap();
        assert_eq!(content, "Copy test content");

        cleanup_path(&src);
        cleanup_path(&dst);
    }

    #[test]
    fn test_fso_movefile() {
        let src = tmp_path("move_src.txt");
        let dst = tmp_path("move_dst.txt");
        cleanup_path(&src);
        cleanup_path(&dst);

        std::fs::write(&src, "Move test content").unwrap();

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     fso.MoveFile \"{}\", \"{}\"",
                    src, dst
                ),
                &mut ctx,
            )
            .unwrap();

        assert!(!std::path::Path::new(&src).exists());
        assert!(std::path::Path::new(&dst).exists());
        let content = std::fs::read_to_string(&dst).unwrap();
        assert_eq!(content, "Move test content");

        cleanup_path(&dst);
    }

    #[test]
    fn test_fso_deletefile() {
        let path = tmp_path("delete_me.txt");
        cleanup_path(&path);
        std::fs::write(&path, "Delete me").unwrap();

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     fso.DeleteFile \"{}\"",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        assert!(!std::path::Path::new(&path).exists());
    }

    #[test]
    fn test_fso_create_delete_folder() {
        let path = tmp_path("test_folder");
        cleanup_path(&path);

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        // Create folder
        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     fso.CreateFolder \"{}\"",
                    path
                ),
                &mut ctx,
            )
            .unwrap();
        assert!(std::path::Path::new(&path).is_dir());

        // Check FolderExists
        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     exists = fso.FolderExists(\"{}\")",
                    path
                ),
                &mut ctx,
            )
            .unwrap();
        assert_eq!(ctx.get_variable("exists"), Some(&VBValue::Boolean(true)));

        // Delete folder
        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     fso.DeleteFolder \"{}\"",
                    path
                ),
                &mut ctx,
            )
            .unwrap();
        assert!(!std::path::Path::new(&path).exists());

        cleanup_path(&path);
    }

    #[test]
    fn test_fso_fileobject_properties() {
        let path = tmp_path("fileobj.txt");
        cleanup_path(&path);
        std::fs::write(&path, "FileObject test").unwrap();

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     Set f = fso.GetFile(\"{}\")\n\
                     name = f.Name\n\
                     size = f.Size",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        // FileObject.Name returns the file name from the path
        let path_obj = std::path::Path::new(&path);
        let expected_name = path_obj.file_name().unwrap().to_str().unwrap().to_string();

        assert_eq!(
            ctx.get_variable("name"),
            Some(&VBValue::String(expected_name))
        );
        assert_eq!(ctx.get_variable("size"), Some(&VBValue::Number(15.0)));

        cleanup_path(&path);
    }

    #[test]
    fn test_fso_folderobject_properties() {
        let path = tmp_path("folderobj");
        cleanup_path(&path);
        std::fs::create_dir_all(&path).unwrap();
        std::fs::write(format!("{}/test.txt", &path), "hello").unwrap();

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     Set fld = fso.GetFolder(\"{}\")\n\
                     fname = fld.Name\n\
                     isRoot = fld.IsRootFolder",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        let path_obj = std::path::Path::new(&path);
        let expected_name = path_obj.file_name().unwrap().to_str().unwrap().to_string();
        assert_eq!(
            ctx.get_variable("fname"),
            Some(&VBValue::String(expected_name))
        );
        assert_eq!(ctx.get_variable("isRoot"), Some(&VBValue::Boolean(false)));

        // Test Files collection
        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     Set fld = fso.GetFolder(\"{}\")\n\
                     files = fld.Files\n\
                     fileCount = 0\n\
                     For Each f In files\n\
                         fileCount = fileCount + 1\n\
                     Next",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        assert_eq!(ctx.get_variable("fileCount"), Some(&VBValue::Number(1.0)));

        cleanup_path(&path);
    }

    #[test]
    fn test_fso_getabsolutepathname() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        interp
            .execute(
                "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                 absPath = fso.GetAbsolutePathName(\"test.txt\")",
                &mut ctx,
            )
            .unwrap();

        let abs = ctx.get_variable("absPath");
        assert!(abs.is_some());
        if let Some(VBValue::String(s)) = abs {
            assert!(s.ends_with("test.txt"));
            assert!(std::path::Path::new(s).is_absolute());
        } else {
            panic!("Expected String");
        }
    }

    #[test]
    fn test_fso_copyfolder() {
        let src = tmp_path("copyfolder_src");
        let dst = tmp_path("copyfolder_dst");
        cleanup_path(&src);
        cleanup_path(&dst);

        std::fs::create_dir_all(&src).unwrap();
        std::fs::write(format!("{}/a.txt", &src), "file a").unwrap();

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     fso.CopyFolder \"{}\", \"{}\", True",
                    src, dst
                ),
                &mut ctx,
            )
            .unwrap();

        assert!(std::path::Path::new(&dst).is_dir());
        assert!(std::path::Path::new(&format!("{}/a.txt", &dst)).exists());

        cleanup_path(&src);
        cleanup_path(&dst);
    }

    #[test]
    fn test_fso_writeblanklines() {
        let path = tmp_path("blanklines.txt");
        cleanup_path(&path);

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     Set ts = fso.CreateTextFile(\"{}\", True)\n\
                     ts.WriteLine \"First\"\n\
                     ts.WriteBlankLines 2\n\
                     ts.WriteLine \"Last\"\n\
                     ts.Close",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        let content = std::fs::read_to_string(&path).unwrap();
        let lines: Vec<&str> = content.lines().collect();
        assert_eq!(lines.len(), 4);
        assert_eq!(lines[0], "First");
        assert_eq!(lines[1], "");
        assert_eq!(lines[2], "");
        assert_eq!(lines[3], "Last");

        cleanup_path(&path);
    }

    #[test]
    fn test_fso_file_delete_method() {
        let path = tmp_path("filedelete.txt");
        cleanup_path(&path);
        std::fs::write(&path, "delete via method").unwrap();

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     Set f = fso.GetFile(\"{}\")\n\
                     f.Delete",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        assert!(!std::path::Path::new(&path).exists());
    }

    #[test]
    fn test_fso_getspecialfolder() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        // TemporaryFolder (2)
        interp
            .execute(
                "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                 tmpFolder = fso.GetSpecialFolder(2)",
                &mut ctx,
            )
            .unwrap();

        let tmp = ctx.get_variable("tmpFolder");
        assert!(tmp.is_some());
        if let Some(VBValue::String(s)) = tmp {
            let p = std::path::Path::new(s);
            assert!(p.is_absolute());
        } else {
            panic!("Expected string path");
        }
    }

    #[test]
    fn test_fso_createobject_invalid_progid() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        let result = interp.execute(
            "Set obj = CreateObject(\"Some.Nonexistent.Object\")",
            &mut ctx,
        );
        assert!(result.is_err());
    }

    #[test]
    fn test_fso_file_exists_false() {
        let path = tmp_path("nonexistent_file.txt");
        cleanup_path(&path);

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     exists = fso.FileExists(\"{}\")",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        assert_eq!(ctx.get_variable("exists"), Some(&VBValue::Boolean(false)));
    }

    #[test]
    fn test_fso_getfile_notfound_error() {
        let path = tmp_path("nonexistent_file_for_get.txt");
        cleanup_path(&path);

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        let result = interp.execute(
            &format!(
                "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                 Set f = fso.GetFile(\"{}\")",
                path
            ),
            &mut ctx,
        );
        assert!(result.is_err());
    }

    #[test]
    fn test_fso_textstream_append() {
        let path = tmp_path("append.txt");
        cleanup_path(&path);
        std::fs::write(&path, "Initial\n").unwrap();

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        // Append mode (8)
        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     Set ts = fso.OpenTextFile(\"{}\", 8, True)\n\
                     ts.WriteLine \"Appended\"\n\
                     ts.Close",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        let content = std::fs::read_to_string(&path).unwrap();
        assert!(content.contains("Appended"));

        cleanup_path(&path);
    }

    #[test]
    fn test_fso_folder_createtextfile() {
        let folder = tmp_path("folder_create_file");
        let file_path = format!("{}/newfile.txt", &folder);
        cleanup_path(&folder);
        std::fs::create_dir_all(&folder).unwrap();

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     Set fld = fso.GetFolder(\"{}\")\n\
                     Set ts = fld.CreateTextFile(\"newfile.txt\", True)\n\
                     ts.WriteLine \"Created via Folder\"\n\
                     ts.Close",
                    folder
                ),
                &mut ctx,
            )
            .unwrap();

        assert!(std::path::Path::new(&file_path).exists());
        let content = std::fs::read_to_string(&file_path).unwrap();
        assert_eq!(content.trim(), "Created via Folder");

        cleanup_path(&folder);
    }

    #[test]
    fn test_fso_file_openastextstream() {
        let path = tmp_path("openas.txt");
        cleanup_path(&path);
        std::fs::write(&path, "OpenAsTextStream content").unwrap();

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     Set f = fso.GetFile(\"{}\")\n\
                     Set ts = f.OpenAsTextStream(1)\n\
                     content = ts.ReadAll()\n\
                     ts.Close",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        assert_eq!(
            ctx.get_variable("content"),
            Some(&VBValue::String("OpenAsTextStream content".to_string()))
        );

        cleanup_path(&path);
    }
