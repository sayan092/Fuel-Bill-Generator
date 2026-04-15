package com.apmm.jcf.common;

import com.olf.openjvs.*;
import com.olf.openjvs.enums.*;
import java.io.*;
import java.util.*;
import java.time.*;
import java.time.temporal.*;
import java.nio.file.*;
import java.nio.file.attribute.FileTime;
import com.olf.openjvs.DBUserTable;
import com.olf.openjvs.ODateTime;
import com.olf.openjvs.Util;

public class FileArchive_Trigger implements IScript {

    private static final String PARAM_CONTEXT_MAIN    = "SHAREPOINT_INTEGRATION";
    private static final String PARAM_CONTEXT_ARCHIVE = "FILESHARE_ARCHIVE";
    private static final String SUBCTX_AUTH           = "AUTH";
    private static final String SUBCTX_CONFIG         = "CONFIG";
    private static final String SUBCTX_ARCHIVE        = "FILE_ARCHIVE";
    private static final String[] PROD_DB_MARKERS     = {
        "FIN104_PROD", "PROD", "PRODUCTION", "FIN104_PRODUCTION"
    };

    @Override
    public void execute(IContainerContext ctx) throws OException {

        Map<String, String> auth    = new LinkedHashMap<>();
        Map<String, String> archive = new LinkedHashMap<>();
        String proxyHost = "";
        String proxyPort = "";

        // -- 1. Load CONFIG params --------------------------------------------
        Table tConfig = Table.tableNew();
        try {
            String sql = "SELECT name, value FROM USER_SCRIPT_PARAMETERS "
                       + "WHERE context     = '" + PARAM_CONTEXT_MAIN + "' "
                       + "AND   sub_context = '" + SUBCTX_CONFIG      + "' "
                       + "AND   LTRIM(RTRIM(ISNULL(value,''))) <> ''";

            if (DBaseTable.execISql(tConfig, sql) != OLF_RETURN_CODE.OLF_RETURN_SUCCEED.toInt())
                throw new OException("Failed to read USER_SCRIPT_PARAMETERS [CONFIG]");

            for (int r = 1; r <= tConfig.getNumRows(); r++) {
                auth.put(tConfig.getString("name", r).trim(),
                         tConfig.getString("value", r).trim());
            }

            OConsole.oprint("CONFIG params loaded: " + tConfig.getNumRows() + " rows\n");

        } finally {
            tConfig.destroy();
        }

        // -- 2. Load AUTH params ----------------------------------------------
        Table tAuth = Table.tableNew();
        try {
            String sql = "SELECT name, value FROM USER_SCRIPT_PARAMETERS "
                       + "WHERE context     = '" + PARAM_CONTEXT_MAIN + "' "
                       + "AND   sub_context = '" + SUBCTX_AUTH        + "' "
                       + "AND   LTRIM(RTRIM(ISNULL(value,''))) <> ''";

            if (DBaseTable.execISql(tAuth, sql) != OLF_RETURN_CODE.OLF_RETURN_SUCCEED.toInt())
                throw new OException("Failed to read USER_SCRIPT_PARAMETERS [AUTH]");

            for (int r = 1; r <= tAuth.getNumRows(); r++) {
                auth.put(tAuth.getString("name", r).trim(),
                         tAuth.getString("value", r).trim());
            }

            OConsole.oprint("AUTH params loaded: " + tAuth.getNumRows() + " rows\n");

        } finally {
            tAuth.destroy();
        }

        // -- 3. Load FILE_ARCHIVE params --------------------------------------
        Table tArchive = Table.tableNew();
        try {
            String sql = "SELECT name, value FROM USER_SCRIPT_PARAMETERS "
                       + "WHERE context     = '" + PARAM_CONTEXT_ARCHIVE + "' "
                       + "AND   sub_context = '" + SUBCTX_ARCHIVE        + "' "
                       + "AND   LTRIM(RTRIM(ISNULL(value,''))) <> ''";

            if (DBaseTable.execISql(tArchive, sql) != OLF_RETURN_CODE.OLF_RETURN_SUCCEED.toInt())
                throw new OException("Failed to read USER_SCRIPT_PARAMETERS [FILE_ARCHIVE]");

            for (int r = 1; r <= tArchive.getNumRows(); r++) {
                archive.put(tArchive.getString("name", r).trim(),
                            tArchive.getString("value", r).trim());
            }

            OConsole.oprint("FILE_ARCHIVE params loaded: " + tArchive.getNumRows() + " rows\n");

        } finally {
            tArchive.destroy();
        }

        // -- 4. Validate mandatory params -------------------------------------
        String[] requiredParams = {
            "TENANT_ID",
            "CLIENT_ID",
            "CLIENT_SECRET",
            "SITE_HOSTNAME",
            "SITE_PATH",
            "AB_ERROR_LOGS_PATH"
        };

        for (String key : requiredParams) {
            if (!auth.containsKey(key) || auth.get(key).isEmpty()) {
                throw new OException("Missing mandatory param: " + key
                                   + " (check CONFIG/AUTH sub_context in USER_SCRIPT_PARAMETERS)");
            }
        }

        if (!archive.containsKey("PS_SCRIPT_PATH") || archive.get("PS_SCRIPT_PATH").isEmpty()) {
            throw new OException("Missing mandatory FILE_ARCHIVE param: PS_SCRIPT_PATH "
                               + "(check FILESHARE_ARCHIVE / FILE_ARCHIVE in USER_SCRIPT_PARAMETERS)");
        }

        String psScriptPath = archive.get("PS_SCRIPT_PATH");

        if (auth.containsKey("PROXY_HOST") && !auth.get("PROXY_HOST").isEmpty())
            proxyHost = auth.get("PROXY_HOST");
        if (auth.containsKey("PROXY_PORT") && !auth.get("PROXY_PORT").isEmpty())
            proxyPort = auth.get("PROXY_PORT");

        OConsole.oprint("PS Script Path : " + psScriptPath + "\n");
        OConsole.oprint("Site Hostname  : " + auth.get("SITE_HOSTNAME") + "\n");
        OConsole.oprint("Site Path      : " + auth.get("SITE_PATH") + "\n");
        OConsole.oprint("Log Path       : " + auth.get("AB_ERROR_LOGS_PATH") + "\n");
        OConsole.oprint("Proxy          : " + (proxyHost.isEmpty() ? "none (direct)"
                                             : proxyHost + ":" + proxyPort) + "\n");

        String currentDbName = getCurrentDatabaseName();
        OConsole.oprint("Database       : " + currentDbName + "\n");

        // -- 5. Load active rules ---------------------------------------------
        Table tRules = Table.tableNew();
        try {
            String sql = "SELECT RULE_ID, RULE_NAME, FROM_PATH, TO_PATH, ACTION_TYPE, "
                       + "       FILE_PATTERN, FILE_AGE_DAYS, RECURSIVE, DELETE_EMPTY_DIRS, "
                       + "       BATCH_SIZE, NO_OF_BATCH_RUN, "
                       + "       ISNULL(PARALLEL_DEGREE, 0) AS PARALLEL_DEGREE "
                       + "FROM USER_FILE_ARCHIVE_CONFIG "
                       + "WHERE ACTIVE = 'Y' "
                       + "ORDER BY RULE_ID";

            if (DBaseTable.execISql(tRules, sql) != OLF_RETURN_CODE.OLF_RETURN_SUCCEED.toInt())
                throw new OException("Failed to read USER_FILE_ARCHIVE_CONFIG");

            int ruleCount = tRules.getNumRows();
            OConsole.oprint("Active rules found: " + ruleCount + "\n");

            if (ruleCount == 0) {
                OConsole.oprint("No active rules found. Exiting.\n");
                return;
            }

            // -- 6. Process each rule -----------------------------------------
            for (int r = 1; r <= ruleCount; r++) {

                int    ruleId          = tRules.getInt   ("RULE_ID",          r);
                String ruleName        = tRules.getString("RULE_NAME",        r).trim();
                String fromPath        = tRules.getString("FROM_PATH",        r).trim();
                String toPath          = tRules.getString("TO_PATH",          r).trim();
                String actionType      = tRules.getString("ACTION_TYPE",      r).trim();
                String filePattern     = tRules.getString("FILE_PATTERN",     r).trim();
                int    fileAgeDays     = tRules.getInt   ("FILE_AGE_DAYS",    r);
                String recursive       = tRules.getString("RECURSIVE",        r).trim();
                String deleteEmptyDirs = tRules.getString("DELETE_EMPTY_DIRS",r).trim();
                int    batchSize       = tRules.getInt   ("BATCH_SIZE",       r);
                int    noOfBatchRun    = tRules.getInt   ("NO_OF_BATCH_RUN",  r);
                int    parallelDegree  = tRules.getInt   ("PARALLEL_DEGREE",  r);

                OConsole.oprint("\n--- Rule [" + ruleId + "] " + ruleName + " ---\n");
                OConsole.oprint("Action         : " + actionType    + "\n");
                OConsole.oprint("From           : " + fromPath      + "\n");
                OConsole.oprint("To             : " + toPath        + "\n");
                OConsole.oprint("Age (days)     : " + fileAgeDays   + "\n");
                OConsole.oprint("Batch          : " + batchSize + " x " + noOfBatchRun + "\n");
                OConsole.oprint("Parallel degree: " + parallelDegree + "\n");

                if (isProductionToPath(toPath) && !isProductionDatabase(currentDbName)) {
                    OConsole.oprint("SKIPPED        : Production ToPath is blocked from non-production database ["
                                  + currentDbName + "]\n");
                    continue;
                }

                // -- Build base PS command ------------------------------------
                List<String> baseCmd = new ArrayList<>();
                baseCmd.add("powershell.exe");
                baseCmd.add("-NonInteractive");
                baseCmd.add("-NoProfile");
                baseCmd.add("-ExecutionPolicy");
                baseCmd.add("Bypass");
                baseCmd.add("-File");
                baseCmd.add(psScriptPath);

                baseCmd.add("-TenantId");        baseCmd.add(auth.get("TENANT_ID"));
                baseCmd.add("-ClientId");        baseCmd.add(auth.get("CLIENT_ID"));
                baseCmd.add("-ClientSecret");    baseCmd.add(auth.get("CLIENT_SECRET"));
                baseCmd.add("-SiteHostname");    baseCmd.add(auth.get("SITE_HOSTNAME"));
                baseCmd.add("-SitePath");        baseCmd.add(auth.get("SITE_PATH"));
                baseCmd.add("-LogBasePath");     baseCmd.add(auth.get("AB_ERROR_LOGS_PATH"));
                baseCmd.add("-RuleId");          baseCmd.add(String.valueOf(ruleId));
                baseCmd.add("-RuleName");        baseCmd.add(ruleName);
                baseCmd.add("-FromPath");        baseCmd.add(fromPath);
                baseCmd.add("-ToPath");          baseCmd.add(toPath);
                baseCmd.add("-ActionType");      baseCmd.add(actionType);
                baseCmd.add("-FilePattern");     baseCmd.add(filePattern.isEmpty() ? "*.*" : filePattern);
                baseCmd.add("-FileAgeDays");     baseCmd.add(String.valueOf(fileAgeDays));
                baseCmd.add("-Recursive");       baseCmd.add(recursive);
                baseCmd.add("-DeleteEmptyDirs"); baseCmd.add(deleteEmptyDirs);
                baseCmd.add("-BatchSize");       baseCmd.add(String.valueOf(batchSize));
                baseCmd.add("-NoOfBatchRun");    baseCmd.add(String.valueOf(noOfBatchRun));

                if (!proxyHost.isEmpty()) {
                    baseCmd.add("-ProxyHost"); baseCmd.add(proxyHost);
                    baseCmd.add("-ProxyPort"); baseCmd.add(proxyPort);
                }

                // -- Execute --------------------------------------------------
                try {
                    boolean success;
                    if (parallelDegree <= 1 || "DELETE_EMPTY_DIRS".equalsIgnoreCase(actionType)) {
                        if (parallelDegree > 1 && "DELETE_EMPTY_DIRS".equalsIgnoreCase(actionType)) {
                            OConsole.oprint("Parallel execution skipped for action ["
                                          + actionType + "] - running single process.\n");
                        }
                        success = runSingleProcess(baseCmd, ruleName, ruleName);
                    } else {
                        success = runParallelProcesses(baseCmd, ruleId, ruleName,
                                                       parallelDegree, fileAgeDays,
                                                       fromPath, filePattern, recursive);
                    }

                    // -- Update LAST_RUN_DT on success ------------------------
                    if (success) {
                        Table updateTable = Util.NULL_TABLE;
                        try {
                            updateTable = Table.tableNew("USER_FILE_ARCHIVE_CONFIG");
                            DBUserTable.structure(updateTable);
                            DBUserTable.load(updateTable);

                            updateTable.sortCol("RULE_ID");
                            int row = updateTable.findInt("RULE_ID", ruleId,
                                                          SEARCH_ENUM.FIRST_IN_GROUP);

                            if (row > 0) {
                                ODateTime now = ODateTime.dtNew();
                                now.setDate(OCalendar.today());
                                now.setTime(Util.timeGetServerTime());
                                updateTable.setDateTime("LAST_RUN_DT", row, now);

                                updateTable.group("RULE_ID");

                                int retval = DBUserTable.update(updateTable);
                                if (retval != OLF_RETURN_CODE.OLF_RETURN_SUCCEED.toInt()) {
                                    OConsole.oprint("WARN: LAST_RUN_DT update failed for rule ["
                                                  + ruleId + "]: "
                                                  + DBUserTable.dbRetrieveErrorInfo(retval,
                                                    "DBUserTable.update() failed") + "\n");
                                } else {
                                    OConsole.oprint("LAST_RUN_DT updated for rule ["
                                                  + ruleId + "]\n");
                                }
                            } else {
                                OConsole.oprint("WARN: Rule [" + ruleId
                                              + "] not found for LAST_RUN_DT update.\n");
                            }
                        } finally {
                            if (Table.isTableValid(updateTable) == 1)
                                updateTable.destroy();
                        }
                    } else {
                        OConsole.oprint("Rule [" + ruleName + "] completed with failures "
                                      + "- LAST_RUN_DT not updated.\n");
                    }

                } catch (Exception e) {
                    OConsole.oprint("EXCEPTION on rule [" + ruleName + "]: "
                                  + e.getMessage() + " - continuing to next rule.\n");
                }
            }

            OConsole.oprint("\n=== All active rules processed. ===\n");

        } finally {
            tRules.destroy();
        }
    }

    // -------------------------------------------------------------------------
    // Run a single PS process and stream output to OConsole
    // -------------------------------------------------------------------------
    private boolean runSingleProcess(List<String> cmd,
                                     String ruleName, String label) throws Exception {
        ProcessBuilder pb = new ProcessBuilder(cmd);
        pb.redirectErrorStream(true);
        Process proc = pb.start();

        BufferedReader reader = new BufferedReader(
            new InputStreamReader(proc.getInputStream()));
        String line;
        while ((line = reader.readLine()) != null) {
            OConsole.oprint("[PS-" + label + "] " + line + "\n");
        }

        int exitCode = proc.waitFor();
        String status = (exitCode == 0) ? "SUCCEEDED" : "FAILED";
        OConsole.oprint("Exit code [" + label + "]: " + exitCode + " (" + status + ")\n");
        return exitCode == 0;
    }

    // -------------------------------------------------------------------------
    // Run N parallel PS processes — auto-shard by FILE_AGE_DAYS window
    // -------------------------------------------------------------------------
    private boolean runParallelProcesses(List<String> baseCmd, int ruleId,
                                         String ruleName, int degree,
                                         int fileAgeDays, String fromPath,
                                         String filePattern, String recursive) throws OException {

        // Auto-shard: split FILE_AGE_DAYS window equally — no manual input needed
        LocalDate cutoff = LocalDate.now().minusDays(fileAgeDays);
        String[] eligibleRange = findEligibleDateRange(fromPath, filePattern, recursive, fileAgeDays);
        if (eligibleRange == null) {
            OConsole.oprint("No eligible files found for parallel sharding. Rule will complete with no work.\n");
            return true;
        }

        LocalDate earliestFrom = LocalDate.parse(eligibleRange[0]);
        LocalDate latestTo     = LocalDate.parse(eligibleRange[1]);
        if (latestTo.isAfter(cutoff)) {
            latestTo = cutoff;
        }

        List<String[]> partitions = buildDatePartitions(earliestFrom.toString(),
                                                        latestTo.toString(), degree);

        OConsole.oprint("Auto-shard range : " + earliestFrom + " -> " + latestTo + "\n");
        OConsole.oprint("Slices           : " + partitions.size() + "\n");

        List<Process> processes = new ArrayList<>();
        List<String>  labels    = new ArrayList<>();
        List<Thread>  threads   = new ArrayList<>();  // track threads for join

        for (int p = 0; p < partitions.size(); p++) {
            String partFrom  = partitions.get(p)[0];
            String partTo    = partitions.get(p)[1];
            String partLabel = ruleName + "_P" + (p + 1);

            List<String> cmd = new ArrayList<>(baseCmd);
            cmd.add("-DateFrom");  cmd.add(partFrom);
            cmd.add("-DateTo");    cmd.add(partTo);
            cmd.add("-PartLabel"); cmd.add(partLabel);

            OConsole.oprint("Partition " + (p + 1) + ": "
                          + partFrom + " -> " + partTo
                          + " [" + partLabel + "]\n");

            try {
                ProcessBuilder pb = new ProcessBuilder(cmd);
                pb.redirectErrorStream(true);
                Process proc = pb.start();

                final String lbl = partLabel;
                Thread t = new Thread(() -> {
                    try {
                        BufferedReader br = new BufferedReader(
                            new InputStreamReader(proc.getInputStream()));
                        String line;
                        while ((line = br.readLine()) != null) {
                            try {
                                OConsole.oprint("[PS-" + lbl + "] " + line + "\n");
                            } catch (OException oe) {
                                System.err.println("[PS-" + lbl + "] " + line);
                            }
                        }
                    } catch (Exception e) {
                        try {
                            OConsole.oprint("Stream error [" + lbl + "]: "
                                          + e.getMessage() + "\n");
                        } catch (OException oe) {
                            System.err.println("Stream error [" + lbl + "]: " + e.getMessage());
                        }
                    }
                });
                t.start();
                threads.add(t);    // save reference for join

                processes.add(proc);
                labels.add(partLabel);

            } catch (Exception e) {
                throw new OException("Failed to spawn PS for partition ["
                                   + partLabel + "]: " + e.getMessage());
            }
        }

        // Wait for all processes to finish
        boolean anyFailed = false;
        for (int p = 0; p < processes.size(); p++) {
            try {
                int exitCode = processes.get(p).waitFor();
                String status = (exitCode == 0) ? "SUCCEEDED" : "FAILED";
                OConsole.oprint("Partition [" + labels.get(p) + "] exit code: "
                              + exitCode + " (" + status + ")\n");
                if (exitCode != 0) anyFailed = true;
            } catch (InterruptedException e) {
                throw new OException("Interrupted waiting for partition ["
                                   + labels.get(p) + "]");
            }
        }

        // Join all output threads — ensures all PS output is fully drained
        // before JVS moves to the next rule
        for (Thread t : threads) {
            try {
                t.join();
            } catch (InterruptedException e) {
                OConsole.oprint("WARN: Output thread interrupted for a partition.\n");
            }
        }

        return !anyFailed;
    }

    private String getCurrentDatabaseName() throws OException {
        Table tDb = Table.tableNew();
        try {
            if (DBaseTable.execISql(tDb, "SELECT DB_NAME() AS DB_NAME") != OLF_RETURN_CODE.OLF_RETURN_SUCCEED.toInt()) {
                throw new OException("Failed to resolve current database name");
            }
            if (tDb.getNumRows() < 1) {
                throw new OException("Current database name query returned no rows");
            }
            return tDb.getString("DB_NAME", 1).trim();
        } finally {
            tDb.destroy();
        }
    }

    private boolean isProductionToPath(String toPath) {
        if (toPath == null) {
            return false;
        }
        String normalized = toPath.trim().toUpperCase(Locale.ROOT);
        return normalized.contains("/PROD/")
            || normalized.contains("\\PROD\\")
            || normalized.startsWith("PROD/")
            || normalized.startsWith("PROD\\")
            || normalized.contains("/PRODUCTION/")
            || normalized.contains("\\PRODUCTION\\")
            || normalized.startsWith("PRODUCTION/")
            || normalized.startsWith("PRODUCTION\\");
    }

    private boolean isProductionDatabase(String dbName) {
        if (dbName == null) {
            return false;
        }
        String normalized = dbName.trim().toUpperCase(Locale.ROOT);
        for (String marker : PROD_DB_MARKERS) {
            if (normalized.equals(marker) || normalized.contains(marker)) {
                return true;
            }
        }
        return false;
    }

    // -------------------------------------------------------------------------
    // Find the actual eligible file date range for a rule
    // -------------------------------------------------------------------------
    private String[] findEligibleDateRange(String fromPath, String filePattern,
                                           String recursive, int fileAgeDays) throws OException {
        try {
            Path start = Paths.get(fromPath);
            if (!Files.exists(start)) {
                return null;
            }

            String pattern = (filePattern == null || filePattern.trim().isEmpty())
                ? "*.*"
                : filePattern.trim();
            PathMatcher matcher = FileSystems.getDefault().getPathMatcher("glob:" + pattern);
            int maxDepth = "Y".equalsIgnoreCase(recursive) ? Integer.MAX_VALUE : 1;

            Instant cutoffInstant = LocalDateTime.now()
                .minusDays(fileAgeDays)
                .atZone(ZoneId.systemDefault())
                .toInstant();

            LocalDate minDate = null;
            LocalDate maxDate = null;

            try (java.util.stream.Stream<Path> stream = Files.walk(start, maxDepth)) {
                Iterator<Path> it = stream.iterator();
                while (it.hasNext()) {
                    Path path = it.next();
                    if (!Files.isRegularFile(path)) {
                        continue;
                    }

                    Path fileName = path.getFileName();
                    if (fileName == null || !matcher.matches(fileName)) {
                        continue;
                    }

                    FileTime lastModified = Files.getLastModifiedTime(path);
                    if (!lastModified.toInstant().isBefore(cutoffInstant)) {
                        continue;
                    }

                    LocalDate fileDate = lastModified.toInstant()
                        .atZone(ZoneId.systemDefault())
                        .toLocalDate();

                    if (minDate == null || fileDate.isBefore(minDate)) {
                        minDate = fileDate;
                    }
                    if (maxDate == null || fileDate.isAfter(maxDate)) {
                        maxDate = fileDate;
                    }
                }
            }

            if (minDate == null || maxDate == null) {
                return null;
            }

            return new String[]{ minDate.toString(), maxDate.toString() };

        } catch (Exception e) {
            throw new OException("Failed to determine eligible file range for ["
                               + fromPath + "]: " + e.getMessage());
        }
    }

    // -------------------------------------------------------------------------
    // Split a date range into N equal slices
    // -------------------------------------------------------------------------
    private List<String[]> buildDatePartitions(String dateFrom, String dateTo, int degree) {

        LocalDate from = LocalDate.parse(dateFrom);
        LocalDate to   = LocalDate.parse(dateTo);

        List<String[]> partitions = new ArrayList<>();
        if (degree <= 1 || !from.isBefore(to)) {
            partitions.add(new String[]{ from.toString(), to.toString() });
            return partitions;
        }

        long totalDaysInclusive = ChronoUnit.DAYS.between(from, to) + 1;
        long baseSliceSize      = totalDaysInclusive / degree;
        long remainder          = totalDaysInclusive % degree;

        LocalDate      sliceStart = from;

        for (int i = 0; i < degree; i++) {
            long sliceSize = baseSliceSize + (i < remainder ? 1 : 0);
            if (sliceSize <= 0) {
                sliceSize = 1;
            }

            LocalDate sliceEnd = sliceStart.plusDays(sliceSize - 1);
            if (sliceEnd.isAfter(to) || i == degree - 1) {
                sliceEnd = to;
            }
            partitions.add(new String[]{ sliceStart.toString(), sliceEnd.toString() });

            if (!sliceEnd.isBefore(to)) {
                break;
            }

            sliceStart = sliceEnd.plusDays(1);
        }
        return partitions;
    }

}
