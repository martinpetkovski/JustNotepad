#pragma once
#include <cstddef>
#include <cstdint>

namespace jnp {
using u8 = uint8_t;
using u16 = uint16_t;
using u32 = uint32_t;
using u64 = uint64_t;
using i64 = int64_t;

constexpr u64 NPOS = ~0ull;
}  // namespace jnp
