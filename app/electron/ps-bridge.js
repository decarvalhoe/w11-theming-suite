/**
 * PowerShell Bridge — Persistent child process for executing w11-theming-suite commands.
 *
 * Spawns a single PowerShell process, loads the suite module, and exposes
 * an async `execute(command)` API that queues commands and resolves with
 * the captured stdout/stderr output.
 */

const { spawn } = require('child_process');
const path = require('path');
const EventEmitter = require('events');

const SENTINEL = '###W11_END###';
const SENTINEL_CMD = `Write-Host '${SENTINEL}'`;

class PowerShellBridge extends EventEmitter {
  constructor() {
    super();
    this.process = null;
    this.ready = false;
    this.queue = [];
    this.current = null;
    this.buffer = '';
    this.errBuffer = '';
    this.stopping = false;
    this.projectRoot = path.resolve(__dirname, '..', '..');
  }

  /** Start the persistent PowerShell process and load the suite module. */
  async start() {
    return new Promise((resolve, reject) => {
      this.stopping = false;

      this.process = spawn('powershell.exe', [
        '-NoProfile',
        '-NoLogo',
        '-ExecutionPolicy', 'Bypass',
        '-Command', '-'
      ], {
        stdio: ['pipe', 'pipe', 'pipe'],
        windowsHide: true
      });

      this.process.stdout.setEncoding('utf8');
      this.process.stderr.setEncoding('utf8');

      this.process.stdout.on('data', (data) => this._onStdout(data));
      this.process.stderr.on('data', (data) => this._onStderr(data));

      this.process.on('close', (code) => {
        this.ready = false;
        this.emit('closed', code);

        // Reject any pending command (including init)
        if (this.current) {
          const { reject: pendingReject } = this.current;
          this.current = null;
          pendingReject(new Error(`PowerShell exited unexpectedly (code ${code})`));
        }

        // Drain the queue — reject all pending commands
        for (const item of this.queue) {
          item.reject(new Error('PowerShell process closed'));
        }
        this.queue = [];

        // Auto-restart after a short delay if unexpected exit (and not stopping)
        if (code !== 0 && !this.stopping) {
          setTimeout(() => this.start().catch(() => {}), 2000);
        }
      });

      this.process.on('error', (err) => {
        this.ready = false;
        this.emit('error', err);
        reject(err);
      });

      // FIX C1: Do NOT double-escape backslashes. PowerShell single-quoted strings
      // treat backslashes as literal characters — no escaping needed.
      const modulePath = path.join(this.projectRoot, 'w11-theming-suite.psd1');
      const initCmd = `Import-Module '${modulePath}' -Force; ${SENTINEL_CMD}\n`;

      this.current = {
        resolve: () => {
          this.ready = true;
          this.current = null;
          this.emit('ready');
          resolve();
          this._processQueue();
        },
        reject,
        stdout: '',
        stderr: ''
      };

      this.buffer = '';
      this.errBuffer = '';
      this.process.stdin.write(initCmd);
    });
  }

  /** Execute a PowerShell command and return its output. */
  execute(command) {
    return new Promise((resolve, reject) => {
      if (this.stopping) {
        reject(new Error('Bridge is stopping'));
        return;
      }
      this.queue.push({ command, resolve, reject });
      if (!this.current && this.ready) {
        this._processQueue();
      }
    });
  }

  /** Stop the PowerShell process gracefully. */
  stop() {
    // FIX I3: Capture process reference and null immediately to prevent
    // auto-restart and double-stop races.
    this.stopping = true;
    const proc = this.process;
    this.process = null;
    this.ready = false;

    if (proc) {
      try {
        proc.stdin.write('exit\n');
      } catch (_) {
        // stdin may already be closed
      }
      setTimeout(() => {
        try { proc.kill(); } catch (_) {}
      }, 3000);
    }
  }

  // ---- Internal ----

  _processQueue() {
    if (this.current || this.queue.length === 0 || !this.ready) return;

    const item = this.queue.shift();
    this.current = {
      resolve: item.resolve,
      reject: item.reject,
      stdout: '',
      stderr: ''
    };
    this.buffer = '';
    this.errBuffer = '';

    // Wrap command so it catches errors and always emits the sentinel
    const wrappedCmd = `try { ${item.command} } catch { Write-Error $_.Exception.Message }; ${SENTINEL_CMD}\n`;
    this.process.stdin.write(wrappedCmd);
  }

  _onStdout(data) {
    this.buffer += data;

    const sentinelIdx = this.buffer.indexOf(SENTINEL);
    if (sentinelIdx !== -1) {
      // Everything before the sentinel is the command output
      const output = this.buffer.substring(0, sentinelIdx).trim();
      this.buffer = this.buffer.substring(sentinelIdx + SENTINEL.length);

      if (this.current) {
        const result = { stdout: output, stderr: this.errBuffer.trim() };
        this.errBuffer = '';
        const cb = this.current.resolve;
        this.current = null;
        cb(result);
        this._processQueue();
      }
    }
  }

  _onStderr(data) {
    this.errBuffer += data;
  }
}

module.exports = PowerShellBridge;
