package local.jakaldocs;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import kr.dogfoot.hwp2hwpx.Hwp2Hwpx;
import kr.dogfoot.hwplib.object.HWPFile;
import kr.dogfoot.hwplib.reader.HWPReader;
import kr.dogfoot.hwpxlib.object.HWPXFile;
import kr.dogfoot.hwpxlib.writer.HWPXWriter;

public class HwpBatchConverterMain {
    public static void main(String[] args) throws Exception {
        Arguments parsed = Arguments.parse(args);
        int failures;

        if (parsed.manifest != null) {
            failures = convertManifest(parsed.manifest, parsed.logPath);
        } else {
            failures = convertSingle(parsed.inputPath, parsed.outputPath, parsed.logPath);
        }

        if (failures > 0) {
            System.exit(1);
        }
    }

    private static int convertSingle(Path inputPath, Path outputPath, Path logPath) throws IOException {
        List<Job> jobs = new ArrayList<Job>(1);
        jobs.add(new Job(inputPath, outputPath));
        return runJobs(jobs, logPath);
    }

    private static int convertManifest(Path manifestPath, Path logPath) throws IOException {
        List<Job> jobs = new ArrayList<Job>();
        try (BufferedReader reader = Files.newBufferedReader(manifestPath, StandardCharsets.UTF_8)) {
            String line;
            while ((line = reader.readLine()) != null) {
                if (line.trim().isEmpty()) {
                    continue;
                }

                String[] parts = line.split("\t", 2);
                if (parts.length != 2) {
                    throw new IllegalArgumentException("Invalid manifest line: " + line);
                }

                jobs.add(new Job(Paths.get(parts[0]), Paths.get(parts[1])));
            }
        }

        return runJobs(jobs, logPath);
    }

    private static int runJobs(List<Job> jobs, Path logPath) throws IOException {
        int failures = 0;
        if (logPath != null && logPath.getParent() != null) {
            Files.createDirectories(logPath.getParent());
        }

        try (BufferedWriter writer = (logPath == null)
                ? null
                : Files.newBufferedWriter(logPath, StandardCharsets.UTF_8)) {
            if (writer != null) {
                writer.write("input\toutput\tstatus\tmessage");
                writer.newLine();
            }

            for (Job job : jobs) {
                String status = "OK";
                String message = "";

                try {
                    convert(job.inputPath, job.outputPath);
                } catch (Exception ex) {
                    failures++;
                    status = "ERROR";
                    message = sanitize(ex.toString());
                }

                if (writer != null) {
                    writer.write(job.inputPath.toString());
                    writer.write('\t');
                    writer.write(job.outputPath.toString());
                    writer.write('\t');
                    writer.write(status);
                    writer.write('\t');
                    writer.write(message);
                    writer.newLine();
                }
            }
        }

        return failures;
    }

    private static void convert(Path inputPath, Path outputPath) throws Exception {
        if (outputPath.getParent() != null) {
            Files.createDirectories(outputPath.getParent());
        }

        HWPFile inputFile = HWPReader.fromFile(inputPath.toString());
        HWPXFile outputFile = Hwp2Hwpx.toHWPX(inputFile);
        HWPXWriter.toFilepath(outputFile, outputPath.toString());
    }

    private static String sanitize(String value) {
        return value.replace('\t', ' ').replace('\r', ' ').replace('\n', ' ');
    }

    private static final class Job {
        private final Path inputPath;
        private final Path outputPath;

        private Job(Path inputPath, Path outputPath) {
            this.inputPath = inputPath;
            this.outputPath = outputPath;
        }
    }

    private static final class Arguments {
        private final Path inputPath;
        private final Path outputPath;
        private final Path manifest;
        private final Path logPath;

        private Arguments(Path inputPath, Path outputPath, Path manifest, Path logPath) {
            this.inputPath = inputPath;
            this.outputPath = outputPath;
            this.manifest = manifest;
            this.logPath = logPath;
        }

        private static Arguments parse(String[] args) {
            Path inputPath = null;
            Path outputPath = null;
            Path manifest = null;
            Path logPath = null;

            for (int index = 0; index < args.length; index++) {
                String current = args[index];
                if ("--input".equals(current)) {
                    inputPath = Paths.get(nextValue(args, ++index, "--input"));
                } else if ("--output".equals(current)) {
                    outputPath = Paths.get(nextValue(args, ++index, "--output"));
                } else if ("--manifest".equals(current)) {
                    manifest = Paths.get(nextValue(args, ++index, "--manifest"));
                } else if ("--log".equals(current)) {
                    logPath = Paths.get(nextValue(args, ++index, "--log"));
                } else {
                    throw new IllegalArgumentException("Unknown argument: " + current);
                }
            }

            if (manifest == null) {
                if (inputPath == null || outputPath == null) {
                    throw new IllegalArgumentException("Use --manifest or provide both --input and --output.");
                }
            } else if (inputPath != null || outputPath != null) {
                throw new IllegalArgumentException("Use either --manifest or --input/--output.");
            }

            return new Arguments(inputPath, outputPath, manifest, logPath);
        }

        private static String nextValue(String[] args, int index, String flag) {
            if (index >= args.length) {
                throw new IllegalArgumentException("Missing value for " + flag);
            }
            return args[index];
        }
    }
}
