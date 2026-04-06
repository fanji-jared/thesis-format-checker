import hashlib
import json
import os
import pickle
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Dict, Optional

from .models import DocumentInfo, ParseResult


@dataclass
class CacheEntry:
    file_path: str
    file_hash: str
    file_size: int
    file_mtime: float
    cached_data: Any
    cached_time: float
    ttl: int = 3600


class DocumentCache:
    DEFAULT_CACHE_DIR = "cache"
    DEFAULT_TTL = 3600
    MAX_CACHE_SIZE = 100 * 1024 * 1024
    
    def __init__(self, cache_dir: Optional[str] = None, ttl: int = DEFAULT_TTL):
        if cache_dir is None:
            cache_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), self.DEFAULT_CACHE_DIR)
        self.cache_dir = Path(cache_dir)
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.ttl = ttl
        self._memory_cache: Dict[str, CacheEntry] = {}
        self._load_disk_cache_index()
    
    def _load_disk_cache_index(self):
        index_file = self.cache_dir / "cache_index.pkl"
        if index_file.exists():
            try:
                with open(index_file, 'rb') as f:
                    self._memory_cache = pickle.load(f)
            except Exception:
                self._memory_cache = {}
    
    def _save_disk_cache_index(self):
        index_file = self.cache_dir / "cache_index.pkl"
        try:
            with open(index_file, 'wb') as f:
                pickle.dump(self._memory_cache, f)
        except Exception:
            pass
    
    def _compute_file_hash(self, file_path: str) -> str:
        hasher = hashlib.md5()
        hasher.update(file_path.encode('utf-8'))
        
        file_stat = os.stat(file_path)
        hasher.update(str(file_stat.st_size).encode('utf-8'))
        hasher.update(str(file_stat.st_mtime).encode('utf-8'))
        
        chunk_size = 65536
        with open(file_path, 'rb') as f:
            f.seek(0, 2)
            file_size = f.tell()
            f.seek(0)
            
            if file_size > 10 * 1024 * 1024:
                for _ in range(3):
                    chunk = f.read(chunk_size)
                    if chunk:
                        hasher.update(chunk)
                    f.seek(file_size // 4, 1)
            else:
                while chunk := f.read(chunk_size):
                    hasher.update(chunk)
        
        return hasher.hexdigest()
    
    def _get_cache_key(self, file_path: str) -> str:
        return hashlib.md5(file_path.encode('utf-8')).hexdigest()
    
    def _is_cache_valid(self, entry: CacheEntry, file_path: str) -> bool:
        if time.time() - entry.cached_time > entry.ttl:
            return False
        
        if not os.path.exists(file_path):
            return False
        
        file_stat = os.stat(file_path)
        if file_stat.st_size != entry.file_size:
            return False
        
        if abs(file_stat.st_mtime - entry.file_mtime) > 1:
            return False
        
        return True
    
    def get(self, file_path: str) -> Optional[DocumentInfo]:
        cache_key = self._get_cache_key(file_path)
        
        if cache_key in self._memory_cache:
            entry = self._memory_cache[cache_key]
            
            if self._is_cache_valid(entry, file_path):
                return entry.cached_data
            else:
                del self._memory_cache[cache_key]
        
        cache_file = self.cache_dir / f"{cache_key}.pkl"
        if cache_file.exists():
            try:
                with open(cache_file, 'rb') as f:
                    entry = pickle.load(f)
                
                if self._is_cache_valid(entry, file_path):
                    self._memory_cache[cache_key] = entry
                    return entry.cached_data
                else:
                    cache_file.unlink()
            except Exception:
                if cache_file.exists():
                    cache_file.unlink()
        
        return None
    
    def set(self, file_path: str, data: DocumentInfo, ttl: Optional[int] = None):
        cache_key = self._get_cache_key(file_path)
        
        file_stat = os.stat(file_path)
        file_hash = self._compute_file_hash(file_path)
        
        entry = CacheEntry(
            file_path=file_path,
            file_hash=file_hash,
            file_size=file_stat.st_size,
            file_mtime=file_stat.st_mtime,
            cached_data=data,
            cached_time=time.time(),
            ttl=ttl or self.ttl
        )
        
        self._memory_cache[cache_key] = entry
        
        cache_file = self.cache_dir / f"{cache_key}.pkl"
        try:
            with open(cache_file, 'wb') as f:
                pickle.dump(entry, f)
            self._save_disk_cache_index()
        except Exception:
            pass
    
    def invalidate(self, file_path: str):
        cache_key = self._get_cache_key(file_path)
        
        if cache_key in self._memory_cache:
            del self._memory_cache[cache_key]
        
        cache_file = self.cache_dir / f"{cache_key}.pkl"
        if cache_file.exists():
            cache_file.unlink()
        
        self._save_disk_cache_index()
    
    def clear(self):
        self._memory_cache.clear()
        
        for cache_file in self.cache_dir.glob("*.pkl"):
            cache_file.unlink()
        
        self._save_disk_cache_index()
    
    def cleanup_expired(self):
        current_time = time.time()
        expired_keys = []
        
        for key, entry in self._memory_cache.items():
            if current_time - entry.cached_time > entry.ttl:
                expired_keys.append(key)
        
        for key in expired_keys:
            del self._memory_cache[key]
            cache_file = self.cache_dir / f"{key}.pkl"
            if cache_file.exists():
                cache_file.unlink()
        
        if expired_keys:
            self._save_disk_cache_index()


def cached_parse(cache: DocumentCache):
    def decorator(func: Callable):
        def wrapper(file_path: str, *args, **kwargs):
            cached_result = cache.get(file_path)
            if cached_result is not None:
                return ParseResult(success=True, document_info=cached_result)
            
            result = func(file_path, *args, **kwargs)
            
            if result.success and result.document_info:
                cache.set(file_path, result.document_info)
            
            return result
        return wrapper
    return decorator
