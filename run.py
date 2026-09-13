if __name__ == '__main__':
    import sys
    import traceback
    from pathlib import Path
    try:
        from oral_registry.app import main
        raise SystemExit(main())
    except Exception:
        if len(sys.argv) > 2 and sys.argv[1] == '--smoke-test':
            Path(sys.argv[2]).with_suffix('.error.txt').write_text(traceback.format_exc(), encoding='utf-8')
            raise SystemExit(1)
        raise
